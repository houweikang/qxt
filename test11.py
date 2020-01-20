import dateutil.parser

start_date = '2019-01-20 19:19:19'

converted_date = dateutil.parser.parse(start_date).date()

print("After converting date is : %s" % converted_date)
