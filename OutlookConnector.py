import datetime as dt
import pandas as pd
import win32com.client

import sys, re

from ftfy import fix_encoding

def print_encoded(s):
    print(fix_encoding(s))

class OutlookConnector:
  olFolderCalendar = 9
  olFolderTodo = 28

  def __init__(self):
    self.outlook = win32com.client.Dispatch('Outlook.Application')
    self.ns = self.outlook.GetNamespace("MAPI")

  def get_calendar(self, start_date, end_date):
      calendar = self.ns.getDefaultFolder(self.olFolderCalendar).Items
      calendar.IncludeRecurrences = True
      calendar.Sort('[Start]')
      formatted_start = start_date.strftime('%m/%d/%Y')
      formatted_end   = end_date.strftime('%m/%d/%Y')
      
      restriction = f"[Start] >= '{formatted_start}' AND [END] <= '{formatted_end}'"
      print(restriction)
      calendar = calendar.Restrict(restriction)
      return calendar

  def get_appointments(self, calendar, subject_kw = None, exclude_subject_kw = None, body_kw = None):
      if subject_kw == None:
          appointments = [app for app in calendar]    
      else:
          appointments = [app for app in calendar if subject_kw in app.subject]
      if exclude_subject_kw != None:
          appointments = [app for app in appointments if exclude_subject_kw not in app.subject]

      print(appointments)
      cal_subject = [app.subject for app in appointments]
      cal_start = [app.start for app in appointments]
      cal_end = [app.end for app in appointments]
      cal_body = [app.body for app in appointments]

      df = pd.DataFrame({'subject': cal_subject,
                        'start': cal_start,
                        'end': cal_end,
                        'body': cal_body})
      return df

  def make_cpd(self, appointments):
      appointments['Date'] = appointments['start']
      print(appointments)
      appointments['Hours'] = (appointments['end'] - appointments['start']).dt.seconds/3600
      appointments.rename(columns={'subject':'Meeting Description'}, inplace = True)
      appointments.drop(['start','end'], axis = 1, inplace = True)
      summary = appointments.groupby('Meeting Description')['Hours'].sum()
      return summary

  def extract_events(self, start_date, end_date):
    calendar = self.get_calendar(start_date, end_date)
    appointments = self.get_appointments(calendar)
    return self.make_cpd(appointments)

  def isMailOrTask(self, item):
    mc = item.MessageClass
    return mc == 'IPM.Task' or mc.startswith('IPM.Note')

  def safeGetIsActive(self, item):
    is_active = False
    try:
      mc = item.MessageClass

      if mc == 'IPM.Task':
        is_active = item.Complete == False
      elif mc.startswith('IPM.Note'):
        is_active = item.IsMarkedAsTask and (item.TaskCompletedDate.year > 3000 or item.FlagStatus == 2)
      else:
        False
    except Exception as e:
      print("Error with email ", item.Subject, str(e))
    finally:
      return is_active

  def getActiveTodos(self):
    todos = self.ns.getDefaultFolder(self.olFolderTodo).Items
    
    print("Getting active todos...")
    active_todos = filter(self.safeGetIsActive, todos)

    return todos
    #for todo in active_todos:
    #  print(todo.Subject)


if __name__ == '__main__':
  begin = dt.datetime.today()-dt.timedelta(days=2)
  end = dt.datetime.today()

  oc = OutlookConnector()
  print("Events: ")
  print(oc.extract_events(begin, end))
  print("---------------------")
  print("Tasks:")
  print(oc.getActiveTodos())