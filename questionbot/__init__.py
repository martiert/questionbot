import asyncio
import aiosparkapi
import json
import pymongo
import os
import re
import xlsxwriter

from spark import Server
import ciscosparkapi



def inviter_is_owner(person, owners):
    for email in person.emails:
        if email in owners:
            return True
    return False


def matches(person, ignores):
    for regex in ignores:
        if regex.search(person.personEmail):
            return True
    return False


class Bot:
    def __init__(self, config, loop):
        self._setup_server(config, loop)
        self._setup_db(config)
        self._api = ciscosparkapi.CiscoSparkAPI(access_token=config['bot']['token'])
        self._owners = config['owners']
        self._room = config['room']
        self._ignore = [re.compile(ignore) for ignore in self._room['ignore']]
        self._questions = config['questions']
        self._users = {}

    async def room_created(self, spark, roomId, person):
        if not roomId in self._room['accepted']:
            return
        if not inviter_is_owner(person, self._owners):
            return
        room = await spark.rooms.get(roomId)
        if not room.type == 'group':
            return

        people = await spark.memberships.list(roomId=roomId)
        async for person in people:
            if not matches(person, self._ignore):
                await self._start_questionere(spark, person.personId)

    async def answer(self, spark, message):
        if not message.personId in self._users.keys():
            return
        user = self._users[message.personId]
        index = user['index']
        question = user['questions'][index]
        question['answer'] = message.text
        user['index'] = index + 1

        if user['index'] == len(user['questions']):
            self._save(user, message.personId)
            await spark.messages.create(toPersonId=message.personId, text='That was all. Thank you for helping out!')
            return

        await self._ask(spark, message.personId)

    async def fetch_answers(self, spark, message):
        if message.personEmail not in self._owners:
            return

        filename = 'answers.xlsx'
        workbook = xlsxwriter.Workbook(filename)

        answer_sheet = workbook.add_worksheet('answers')
        index = 0
        for question in self._questions:
            answer_sheet.write(0, index, question['question'])
            index += 1

        row = 1
        for answer in self._db.find({}):
            index = 0
            for question in answer['questions']:
                answer_sheet.write(row, index, question['answer'])
                index += 1
            row += 1
        workbook.close()

        self._api.messages.create(toPersonId=message.personId, files=[filename])
        os.remove(filename)

    async def _start_questionere(self, spark, personId):
        self._users[personId] = {
                'questions': self._questions[:],
                'index': 0,
        }
        await spark.messages.create(
                toPersonId=personId,
                text='Hello Designers. To shape our TDG new hire training, we would love it if you would take the time to answer this short survey'
        )
        await self._ask(spark, personId)

    async def _ask(self, spark, personId):
        user = self._users[personId]
        index = user['index']

        question = user['questions'][index]['question']
        await spark.messages.create(toPersonId=personId, text=question)

    def _save(self, user, personId):
        self._db.insert_one(user)
        del self._users[personId]

    def _setup_server(self, config, loop):
        self._server = Server(config['bot'], loop)

        self._server.listen(
                '^answers$',
                self.fetch_answers)
        self._server.roomcreation(self.room_created)
        self._server.default_message(self.answer)

        loop.run_until_complete(self._server.setup(loop))

    def _setup_db(self, config):
        conf = config['database']
        client = pymongo.MongoClient('mongodb://127.0.0.1')
        self._db = client[conf['collection']][conf['database']]

    def run(self, loop):
        print('======== Bot Ready ========')
        try:
            loop.run_forever()
        except KeyboardInterrupt:
            pass
        except:
            print(sys.exc_info())
        finally:
            loop.run_until_complete(self._server.cleanup())
