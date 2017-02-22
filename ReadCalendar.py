import win32com
import  datetime
from dateutil.relativedelta import relativedelta

Outlook = win32com.client.Dispatch("Outlook.Application")
ns = Outlook.GetNamespace("MAPI")

appointments = namespace.GetDefaultFolder(9).Items 
# TODO: Need to figure out howto get the shared calendar instead Default [9] 
# (I have placed the shared folder into a separate folder - don't know if it matters)
# I would just like the user to select which calendar to execute on
appointments.Sort("[Start]")
appointments.IncludeRecurrences = "True"
begin = date.today().strftime("%m%d%Y")
end = (date.today() + relativedelta( months = 3 )).strftime("%m%d%Y")
appointments = appointments.Restrict("[Start] >= '" +begin+ "' AND [END] >= '" +end+ "'")


# Get the AppointmentItem objects
# http://msdn.microsoft.com/en-us/library/office/aa210899(v=office.11).aspx

# Restrict to items in the next 30 days (using Python 3.3 - might be slightly different for 2.7)
begin = datetime.date.today()
end = begin + datetime.timedelta(days = 30);
restriction = "[Start] >= '" + begin.strftime("%m/%d/%Y") + "' AND [End] <= '" +end.strftime("%m/%d/%Y") + "'"
restrictedItems = appointments.Restrict(restriction)

# Iterate through restricted AppointmentItems and print them
for appointmentItem in restrictedItems:
    print("{0} Start: {1}, End: {2}, Organizer: {3}".format(
          appointmentItem.Subject, appointmentItem.Start, 
          appointmentItem.End, appointmentItem.Organizer))