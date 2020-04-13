
from datetime import date
import datetime
import re

date_pattern = '{{ADJUST:${Contact.CreatedDate}(10,0,0)}}'
adjust_pattern  = re.findall("\\{{ADJUST\\:(.*?)\\}}", date_pattern)
matched_patterns = re.findall("\\$\\{(.*?)\\}", adjust_pattern[0])
format_type = re.findall("\\((.*?)\\)",adjust_pattern[0])[0]
format_type = format_type.split(',')
date_value = '2020-03-16T05:11:04.000+0000'
custom_date_value = '2020-06-17'
separate_date = date_value.split('-')
datefield = separate_date[2][:2]
print("printpattern-->{}".format(adjust_pattern))
print("matched_patterns-->{}".format(matched_patterns))
print("format_type-->{}".format(format_type))
value = date(int(separate_date[0]), int(separate_date[1]), int(datefield))
result = value + datetime.timedelta(int(format_type[1])*365/12)
adding_days = value + datetime.timedelta(days=int(format_type[0]))
adding_years = value + datetime.timedelta(int(format_type[2])*365)
print(value)
print(result)
print(adding_days)
print(adding_years)