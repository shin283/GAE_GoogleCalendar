# -*- coding: utf-8 -*-

import webapp2
import os
import yaml

from apiclient.discovery import build
from oauth2client.appengine import OAuth2Decorator
from google.appengine.ext import ndb

# api.yamlから読み取り
api_key = yaml.safe_load(open('api.yaml').read().decode('utf-8'))

decorator = OAuth2Decorator(
        client_id=api_key['client_id'],
        client_secret=api_key['client_secret'],
        scope='https://www.googleapis.com/auth/calendar',
        )

# オブジェクトを生成。build() は JSON のサービス定義ファイルからメソッドを構築するためのもの。
service = build('calendar', 'v3')

# add start
DEFAULT_GUESTBOOK_NAME = 'default_guestbook'

def guestbook_key(guestbook_name=DEFAULT_GUESTBOOK_NAME):
    """Constructs a Datastore key for a Guestbook entity with guestbook_name."""
    return ndb.Key('CalendarRead', guestbook_name)

class Greeting(ndb.Model):
    """Models an individual Guestbook entry with author, content, and date."""
    # author = ndb.UserProperty()
    id = ndb.StringProperty(indexed=False)
    summary = ndb.StringProperty(indexed=False)
    description = ndb.StringProperty(indexed=False)
    location = ndb.StringProperty(indexed=False)
    start = ndb.StringProperty(indexed=False)
    end = ndb.StringProperty(indexed=False)
    # start = ndb.DateTimeProperty(auto_now_add=True)
    # end = ndb.DateTimeProperty(auto_now_add=True)
    # date = ndb.DateTimeProperty(auto_now_add=True)
# add end

class IndexHandler(webapp2.RequestHandler):
    def get(self):
        MAIN_PAGE_HTML = """\
        <!DOCTYPE html>
          <body>
            <h1>Delete</h1>
            <p>件名、期間を指定して予定削除できる機能。</p>
            <p>たとえば、2/9〜2/11までの間で、会議名「会議予約」の予定を削除したいときは↓のように入力。</p>
            <form action="/delete" method="get">
              <div>件名:<input type="text" name="summary" value="【会議予約" style="width:300px;"></text></div>
              <div>From:<input type="text" name="timeMin" value="2014-02-09T00:00:00+09:00" style="width:300px;"></text></div>
              <div>To:<input type="text" name="timeMax" value="2014-02-11T00:00:00+09:00" style="width:300px;"></text></div>
              <div><input type="submit" value="delete"></div>
            </form>
            <br>
            <br>
            <br>
            <br>
            - - - - - - -
            <br>
            <h1>Create</h1>
            <p>Google Calendarに予定を登録できる機能。</p>
            <p>Calendarから登録するのと同じ。</p>
            <form action="/create" method="get">
              <div>件名:<input type="text" name="summary" value="無題の予定" style="width:300px;"></text></div>
              <div>説明:<input type="text" name="description" value="会議の説明" style="width:300px;"></text></div>
              <div>場所:<input type="text" name="location" value="汐留" style="width:300px;"></text></div>
              <div>開始時間:<input type="text" name="startdateTime" value="2014-02-10T10:00:00.000" style="width:300px;"></text></div>
              <div>終了時間:<input type="text" name="enddateTime" value="2014-02-10T11:00:00.000" style="width:300px;"></text></div>
              <div>タイムゾーン:<input type="text" name="timeZone" value="Asia/Tokyo" style="width:300px;"></text></div>
              <div><input type="submit" value="add"></div>
            </form>
            <br>
            - - - - - - -
            <br>
            <h1>Read</h1>
            <p>期間を指定して予定を抽出できる機能。</p>
            <form action="/read" method="get">
              <div>件名:<input type="text" name="summary" value="会議予約" style="width:300px;"></text></div>
              <div>From:<input type="text" name="timeMin" value="2014-02-09T00:00:00+09:00" style="width:300px;"></text></div>
              <div>To:<input type="text" name="timeMax" value="2014-02-11T00:00:00+09:00" style="width:300px;"></text></div>
              <div><input type="submit" value="read"></div>
            </form>
            <br>
            - - - - - - -
            <br>
            <h1>Update</h1>
            <p>404 Not Found.</p>
          </body>
        </html>
        """
        self.response.write(MAIN_PAGE_HTML)

class CalendarCreateHandler(webapp2.RequestHandler):

    @decorator.oauth_required
    def get(self):

        # HTMLから値を取得
        summary = self.request.get('summary')
        description = self.request.get('description')
        location = self.request.get('location')
        startdateTime = self.request.get('startdateTime')
        enddateTime = self.request.get('enddateTime')
        timeZone = self.request.get('timeZone')

        # 登録するeventのデータを作成
        event = {
            'summary': summary,
            'description': description,
            'location': location,
            'start': {
                'dateTime': startdateTime,
                'timeZone': timeZone
            },
            'end': {
                'dateTime': enddateTime,
                'timeZone': timeZone
            },
        }

        # insertメソッドで登録
        request = service.events().insert(
            calendarId='primary',
            body=event)

        http = decorator.http()
        request.execute(http=http)

        # Topに戻るLink
        ToTop = """\
        <html>
          <body>
            <a href="/">To Top
          </body>
        </html>
        """

        self.response.write(ToTop)

class CalendarReadHandler(webapp2.RequestHandler):

    @decorator.oauth_required
    def get(self):

        # HTMLから値を取得
        timeMax = self.request.get('timeMax')
        timeMin = self.request.get('timeMin')

        page_token = None
        while True:
            request = service.events().list(
                calendarId = 'primary',
                timeMax = timeMax,
                timeMin = timeMin,
                pageToken=page_token)
            http = decorator.http()
            events = request.execute(http=http)

            for event in events['items']:
                # 削除対象の件名を判断し、削除
                greeting = Greeting(parent=guestbook_key(event['id']))
                greeting.id = event['id']
                greeting.summary = event['summary']
                try:
                  greeting.description = event['description']
                except:
                  greeting.location = 'None'
                try:
                  greeting.location = event['location']
                except:
                  greeting.location = 'None'
                greeting.start = event['start']['dateTime']
                greeting.end = event['end']['dateTime']
                greeting.put()
            page_token = events.get('nextPageToken')
            if not page_token:
                break

        # Topに戻るLink
        ToTop = """\
        <html>
          <body>
            <a href="/">To Top
          </body>
        </html>
        """

        self.response.write(ToTop)

class CalendarDeleteHandler(webapp2.RequestHandler):

    @decorator.oauth_required
    def get(self):

        # HTMLから値を取得
        summary = self.request.get('summary')
        timeMax = self.request.get('timeMax')
        timeMin = self.request.get('timeMin')

        page_token = None
        while True:
            request = service.events().list(
                calendarId = 'primary',
                timeMax = timeMax,
                timeMin = timeMin,
                pageToken=page_token)
            http = decorator.http()
            events = request.execute(http=http)
            for event in events['items']:
                # 削除対象の件名を判断し、削除
                if event['summary'] == summary:
                    service.events().delete(calendarId='primary', eventId=event['id']).execute(http=http)
            page_token = events.get('nextPageToken')
            if not page_token:
                break

        # Topに戻るLink
        ToTop = """\
        <html>
          <body>
            <a href="/">To Top
          </body>
        </html>
        """

        self.response.write(ToTop)


debug = os.environ.get('SERVER_SOFTWARE', '').startswith('Dev')

app = webapp2.WSGIApplication([
                               ('/', IndexHandler),
                               ('/create', CalendarCreateHandler),
                               ('/read', CalendarReadHandler),
                               ('/delete', CalendarDeleteHandler),
                               (decorator.callback_path, decorator.callback_handler()),
                              ],
                              debug=debug)
