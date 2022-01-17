import datetime as dt
import win32com.client


from ftfy import fix_encoding


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

    def todos_to_markdown(self, todos):
        result = ""

        for todo in todos:
            try:
                result += f"[ ] {fix_encoding(todo.Subject)}\n"
            except:
                pass

        return result

    def todays_agenda_as_markdown(self):
        calendar=self.get_events(dt.date.today(),
                                   dt.date.today() + dt.timedelta(days=1))

        return self.events_to_markdown(calendar)

    def get_active_todos(self):
        todos = self.ns.getDefaultFolder(self.olFolderTodo).Items
        
        # Task with [Complete] = False or Note with FlagStatus != 1
        MessageClass = "http://schemas.microsoft.com/mapi/proptag/0x001a001e"
        Complete = "http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/811c000b"
        FlagStatus = "http://schemas.microsoft.com/mapi/proptag/0x10900003"

        restriction = f"@SQL=(({MessageClass}='IPM.Task' AND {Complete}=0) OR ({MessageClass}='IPM.Note' AND (NOT {FlagStatus}=1)))"
        
        active_todos = todos.Restrict(f'{restriction}')

        return active_todos

    def active_todos_as_markdown(self):
        return self.todos_to_markdown(self.get_active_todos())

if __name__ == '__main__':
    import datetime as dt
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("-e", "--print_events", action="store_true", 
                        help="Print markdown representation of today's events")
    parser.add_argument("-t", "--print_tasks",  action="store_true",
                        help="Print markdown representation of open tasks")

    args = parser.parse_args()

    oc = OutlookConnector()

    if args.print_events:
        print(oc.todays_agenda_as_markdown())

    if args.print_tasks:
        print(oc.active_todos_as_markdown())