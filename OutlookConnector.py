import datetime as dt
import pandas as pd
import win32com.client

import sys
import re

from ftfy import fix_encoding


def print_encoded(s):
    print(fix_encoding(s))


class OutlookConnector:
    olFolderCalendar = 9
    olFolderTodo = 28

    def __init__(self):
        self.outlook = win32com.client.Dispatch('Outlook.Application')
        self.ns = self.outlook.GetNamespace("MAPI")
        self.accounts = self.outlook.Session.Accounts

    def get_calendar_for_account(self, account):
        recipient = self.outlook.Session.createRecipient(account.DisplayName)
        calendar = self.outlook.Session.GetSharedDefaultFolder(
            recipient, self.olFolderCalendar).Items

        return calendar

    def get_events_for_account(self, account, start_date, end_date):
        calendar = self.get_calendar_for_account(account)
        calendar.IncludeRecurrences = True
        calendar.Sort('[Start]')
        formatted_start = start_date.strftime('%m/%d/%Y')
        formatted_end = end_date.strftime('%m/%d/%Y')

        restriction = f"[Start] >= '{formatted_start}' AND [Start] < '{formatted_end}'"

        calendar = calendar.Restrict(restriction)
        return calendar

    def get_events(self, start_date, end_date):
        events_for_all = [
            self.get_events_for_account(
                account,
                start_date,
                end_date) for account in self.accounts
        ]

        return sorted(
            [event for events in events_for_all for event in events],
            key=lambda e: e.Start
        )

    def events_to_markdown(self, appointments):
        result=""

        for app in appointments:
            result += f"### {app.Start.strftime('%H:%M')}-{app.END.strftime('%H:%M')}: {app.Subject}\n\n"

        return result

    def todays_agenda_as_markdown(self):
        calendar=self.get_events(dt.date.today(),
                                   dt.date.today() + dt.timedelta(days=1))

        return self.events_to_markdown(calendar)

    def isMailOrTask(self, item):
        mc=item.MessageClass
        return mc == 'IPM.Task' or mc.startswith('IPM.Note')

    def safeGetIsActive(self, item):
        is_active=False
        try:
            mc=item.MessageClass

            if mc == 'IPM.Task':
                is_active=item.Complete == False
            elif mc.startswith('IPM.Note'):
                is_active=item.IsMarkedAsTask and (
                    item.TaskCompletedDate.year > 3000 or item.FlagStatus == 2)
            else:
                False
        except Exception as e:
            print("Error with email ", item.Subject, str(e))
        finally:
            return is_active

    def getActiveTodos(self):
        todos = self.ns.getDefaultFolder(self.olFolderTodo).Items

        print("Getting active todos...")
        active_todos = list(filter(self.safeGetIsActive, todos))

        return active_todos
        # for todo in active_todos:
        #  print(todo.Subject)


if __name__ == '__main__':
    from OutlookConnector import OutlookConnector
    import datetime as dt

    oc = OutlookConnector()
    events = oc.get_events(
        dt.date.today(),
        dt.date.today() +
        dt.timedelta(
            days=1))
    print("Events: ")
    print(
        oc.get_events(
            dt.date.today(),
            dt.date.today() +
            dt.timedelta(
                days=1)))
    print("---------------------")
    print("Tasks:")
    print(oc.getActiveTodos())

    begin = dt.datetime.today() - dt.timedelta(days=2)
    end = dt.datetime.today()

    oc = OutlookConnector()
