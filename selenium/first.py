
from imports import *

db_file = "db.csv"
check_db_errors(db_file)
# open_mssql()
# connectServer()
driver = init_driver()
driver = open_case_queue_sf(driver)
driver = handling_new_ones(driver,db_file)

