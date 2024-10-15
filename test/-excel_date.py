def xls_date_to_ordinal(xls_date):
    xls_date = xlrd.xldate_as_tuple(xls_date, 0)
    xls_date = datetime(*xls_date).date()
    xls_date = datetime.toordinal(xls_date)
    return xls_date

