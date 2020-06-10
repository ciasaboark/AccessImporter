import csv, time, json, logging

import pyodbc
import hashlib

#The imported data will be considered stale after 4 days
MAX_AGE_MS = 1000 * 60 * 60 * 24 * 4

logger = logging.getLogger('watcher.importer')

class Importer:
    def __init__(self, filename,):
        self.filename = filename

    def md5(self, filename):
        hash_md5 = hashlib.md5()
        with open (filename, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()

    def begin_import(self):
        logger.info("Beginning import of file {0}".format(self.led))
        pass
        # mariadb_connection = mariadb.connect(user='returns', password='', database='returns')
        # cursor = mariadb_connection.cursor(dictionary=True)
        # md5 = ""

        # #check if this file was already imported, if so we can just return
        # try:
        #     md5 = self.md5(self.filename)
        #     logger.debug("MD5 sum for file '{}': {}".format(self.filename, md5))

        #     #if
        #     query = ("SELECT * FROM meta")
        #     cursor.execute(query)
        #     row = cursor.fetchone()


        #     if row is None:
        #         logger.debug("No meta data recorded for previous file import, beginning import.")
        #     else:
        #         #we do have some meta data.  Check if a md5 sum was stored
        #         status = row
        #         if 'md5' not in status:
        #             logger.debug("Previous meta data does not have md5 sum recorded, beginning import.")
        #         else:
        #             #we do have a previous md5.  Make sure this file does not have the same hash
        #             old_md5 = status['md5']
        #             if md5 == old_md5:
        #                 logger.debug("File '{}' has already been imported.  Skipping import.".format(self.filename))
        #                 return

        # except Exception as e:
        #     logger.debug("Unable to read md5sum of file '{}'.  Skipping import".format(self.filename))
        #     raise  e

        # if self.led is not None:
        #     self.led.blink(on_time=.2, off_time=.2)
        # try:
        #     products = {}
        #     locations = {}
        #     upcs = {}

        #     with open(self.filename) as csv_file:
        #         csv_reader = csv.reader(csv_file, delimiter=',')
        #         line_count = 0
        #         for row in csv_reader:
        #             if line_count == 0:
        #                 logger.debug("Column names are {}".format(row))
        #                 line_count += 1
        #             else:
        #                 dc_id = row[0]
        #                 whse_id = row[1]
        #                 loc_id = row[2]
        #                 prod_id = row[3]
        #                 case_cost = row[4]
        #                 lcus_id = row[5]
        #                 upc_id = row[6]
        #                 lcat_id = row[7]
        #                 x_coord = row[8]
        #                 y_coord = row[9]
        #                 z_coord = row[10]
        #                 lev = row[11]
        #                 description = row[12]
        #                 vend_prod_no = row[13]
        #                 ldes_id = row[14]
        #                 sel_pos_hgt = row[15]
        #                 rsv_pos_hgt = row[16]
        #                 stk_pos_wid = row[17]
        #                 stk_pos_dep = row[17]

        #                 p = {
        #                     'dc_id': dc_id,
        #                     'whse_id': whse_id,
        #                     'prod_id': prod_id,
        #                     'case_cost': case_cost,
        #                     'description': description,
        #                     'vend_prod_no': vend_prod_no,
        #                 }

        #                 l = {
        #                     'loc_id': loc_id,
        #                     'lcus_id': lcus_id,
        #                     'lcat_id': lcat_id,
        #                     'x_coord': x_coord,
        #                     'y_coord': y_coord,
        #                     'z_coord': z_coord,
        #                     'lev': lev,
        #                     'ldes_id': ldes_id,
        #                     'sel_pos_hgt': sel_pos_hgt,
        #                     'rsv_pos_hgt': rsv_pos_hgt,
        #                     'stk_pos_wid': stk_pos_wid,
        #                     'stk_pos_dep': stk_pos_dep,
        #                     'prod_id': prod_id
        #                 }

        #                 u = {
        #                     'upc_id': upc_id,
        #                     'prod_id': prod_id
        #                 }

        #                 if prod_id not in products.keys():
        #                     products[prod_id] = p

        #                 if loc_id+prod_id not in locations.keys():
        #                     locations[loc_id+prod_id] = l

        #                 if upc_id+prod_id not in upcs.keys():
        #                     upcs[upc_id+prod_id] = u

        #         logger.debug("Read {} items in {} locations with {} upcs from product listing".format(len(products), len(locations), len(upcs)))


        #         #Update the product listing catalog
        #         if self.led is not None:
        #             self.led.blink(on_time=.8, off_time=.8)

        #         #Insert the data from the three dictionaries into the databases
        #         #Clear any previous data
        #         cursor = mariadb_connection.cursor()
        #         cursor.execute("TRUNCATE TABLE products")
        #         cursor.execute("TRUNCATE TABLE locations")
        #         cursor.execute("TRUNCATE TABLE upcs")

        #         add_product_sql = ("INSERT INTO products "
        #                        "(dc_id, whse_id, prod_id, case_cost, description, vend_prod_no)"
        #                        "VALUES (%s, %s, %s, %s, %s, %s)")
        #         add_location_sql = ("INSERT INTO locations "
        #                             "(loc_id, lcus_id, lcat_id, x_coord, y_coord, z_coord, lev, ldes_id, sel_pos_hgt, rsv_pos_hgt, stk_pos_wid, stk_pos_dep, prod_id)"
        #                             "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)")
        #         add_upc_sql = ("INSERT INTO upcs "
        #                            "(upc_id, prod_id)"
        #                            "VALUES (%s, %s)")

        #         logger.debug("Inserting products...")
        #         for p in products.values():
        #             data = (
        #                 p['dc_id'],
        #                 p['whse_id'],
        #                 p['prod_id'],
        #                 p['case_cost'],
        #                 p['description'],
        #                 p['vend_prod_no']
        #             )
        #             cursor.execute(add_product_sql, data)

        #         logger.debug("Done")


        #         logger.debug("Inserting locations...")
        #         for l in locations.values():
        #             data = (
        #                 l['loc_id'],
        #                 l['lcus_id'],
        #                 l['lcat_id'],
        #                 l['x_coord'],
        #                 l['y_coord'],
        #                 l['z_coord'],
        #                 l['lev'],
        #                 l['ldes_id'],
        #                 l['sel_pos_hgt'],
        #                 l['rsv_pos_hgt'],
        #                 l['stk_pos_wid'],
        #                 l['stk_pos_dep'],
        #                 l['prod_id']
        #             )
        #             cursor.execute(add_location_sql, data)
        #         logger.debug("Done")


        #         logger.debug("Inserting upcs...")
        #         for u in upcs.values():
        #             data = (
        #                 u['upc_id'],
        #                 u['prod_id']
        #             )
        #             cursor.execute(add_upc_sql, data)
        #         logger.debug("Done")

        #         #Update the meta status catalog so we can monitor the product catalog age
        #         logger.debug("Updating status table")
        #         timestamp = int(time.time()*1000)
        #         expiration = timestamp + MAX_AGE_MS

        #         add_meta_sql = ("INSERT INTO meta "
        #                        "(num_records, last_db_sync, max_db_age, md5)"
        #                        "VALUES (%s, %s, %s, %s)")
        #         status = (
        #             len(products),
        #             timestamp,
        #             expiration,
        #             md5
        #         )

        #         cursor.execute("TRUNCATE TABLE meta")
        #         cursor.execute(add_meta_sql, status)
        #         logger.debug("Done")

        #         logger.debug("Committing changes...")
        #         mariadb_connection.commit()
        #         logger.debug("Done")

        #         #update the display
        #         try:
        #             logger.debug("Updating display")
        #             import display
        #             display.update_display()
        #         except Exception as e:
        #             logger.error("Error updating display after data import")

        # except Exception as e:
        #     logger.error("Error reading product listing file.")
        #     raise e



def to_dict(obj):
    return json.loads(json.dumps(obj, default=lambda o: o.__dict__))