import fitz
import pandas as pd
import re
import glob
import os
from faker import Faker
from datetime import datetime, timedelta
from random import randint

fake = Faker()

file_path = os.path.dirname(__file__)
os.chdir(file_path)


class PurchaseOrderReader:
    def __init__(self, pdf):
        self.pdf = pdf
        self.pymupdf_text = ""
        self.orderlines_list = []
        self.purchaseheader_extracted_dict = dict(
            zip(
                ["ordernumber", "order_date", "delivery_terms", "payment_terms"],
                [[] for _ in range(4)],
            )
        )
        self.purchaserow_extracted_dict = dict(
            zip(
                [
                    "ordernumber",
                    "orderline",
                    "item",
                    "delivery_date",
                    "quantity",
                    "price_per_piece",
                    "delivery_terms",
                    "payment_terms",
                ],
                [[] for _ in range(8)],
            )
        )
        self.process_mining_dict = dict(
            zip(
                ["case", "timestamp", "activity"],
                [[] for _ in range(3)],
            )
        )
        self.process_mining_dict = dict(
            zip(
                ["case", "timestamp", "activity"],
                [[] for _ in range(3)],
            )
        )
        self.general_processes_dictonary = {
            "process_start": [
                "Extracted PDF from mail",
                "Extracted data from PDF",
                "Extracted ordernumber from data",
            ],
            "ordernumber": [
                "Ordernumber was found in ERP-system",
                "Ordernumber was not found in ERP-system",
                "Mail was sent to the purchasing department",
            ],
            "delivery_terms": [
                "Delivery terms matched with masterdata",
                "Delivery terms didn't match with masterdata",
                "Mail was sent to the purchasing department",
            ],
            "payment_terms": [
                "Payment terms matched with masterdata",
                "Payment terms didn't match with masterdata",
                "Mail was sent to the purchasing department",
            ],
            "item": [
                "Items matched with masterdata",
                "Item didn't match with masterdata",
                "Mail was sent to the purchasing department",
            ],
            "delivery_date": [
                "Delivery date extracted from data",
                "Delivery date filled in the ERP-system",
                "Delivery date was not found in data",
                "Mail was sent to the purchasing department",
            ],
            "quantity": [
                "Quantity matched with masterdata",
                "Quantity didn't match with masterdata",
                "Mail was sent to the purchasing department",
            ],
            "price_per_piece": [
                "Price per piece matched with masterdata",
                "Price per piece didn't match with masterdata",
                "Mail was sent to the purchasing department",
            ],
        }

        self.general_information_patterns = [
            r"(Date:\D[n]\d{1,2}-\d{1,2}-\d{4})",
            r"(Payment terms:(.*?)days.)",
            r"(Delivery terms:(.*?)TAX)",
        ]

    def import_and_merge_snappies(self):
        self.purchaseheader_database = pd.read_parquet(
            r"fake_confirmations\snappy\total_purchaseheader_df.snappy.parquet"
        )
        self.purchaserow_database = pd.read_parquet(r"fake_confirmations\snappy\total_purchaserow_df.snappy.parquet")
        self.purchase_total_database = self.purchaserow_database.merge(
            self.purchaseheader_database[["ordernumber", "delivery_terms", "payment_terms"]],
            on=["ordernumber"],
            how="left",
        )
        return self.purchase_total_database

    def pymupdf_data_extractor(self):
        with fitz.open(self.pdf) as doc:
            for page in doc:
                self.pymupdf_text = "{}\n{}".format(self.pymupdf_text, repr(page.get_text()))
        return self.pymupdf_text

    def extract_ordernumber_and_comparing_ordernumber_with_database(self):
        ordernr_pattern = r"(Your reference:\D[n]92\d{1,6})"
        matches = re.search(ordernr_pattern, self.pymupdf_text)
        self.ordernumber = str(matches.group().replace("Your reference:\\n", "").replace("'", "").strip())
        self.purchase_total_database_filtered = self.purchase_total_database[
            self.purchase_total_database["ordernumber"] == self.ordernumber
        ]
        #  print(self.purchase_total_database_filtered)
        if len(self.purchase_total_database_filtered.index) == 0:
            self.process_continue = False
            self.start_rpa_process(self.ordernumber)
            for key in self.general_processes_dictonary["ordernumber"][1:]:
                self.time = self.time + timedelta(seconds=self.sleep_for_a_while())
                self.process_mining_dict["case"].append(self.ordernumber)
                self.process_mining_dict["timestamp"].append(self.time)
                self.process_mining_dict["activity"].append(key)
        else:
            self.process_continue = True
        return self.process_mining_dict, self.process_continue, self.ordernumber

    def extract_general_information(self):
        for idx, key in enumerate(self.general_information_patterns):
            matches = re.search(key, self.pymupdf_text)
            if idx == 0:
                self.order_date = matches.group().replace("Date:\\n", "").replace("'", "")
            elif idx == 1:
                self.payment_terms = matches.group().replace("Payment terms:\\n", "").replace("'", "")
            elif idx == 2:
                self.delivery_terms = (
                    matches.group().replace("Delivery terms:\\n", "").replace("\\nTAX", "").replace("'", "")
                )
        keys = ["ordernumber", "order_date", "delivery_terms", "payment_terms"]
        values = [
            self.ordernumber,
            self.order_date,
            self.delivery_terms,
            self.payment_terms,
        ]
        self.purchaseheader_extracted_dict = dict(zip(keys, values))
        return self.purchaseheader_extracted_dict, self.ordernumber

    def transforming_orderlines_raw(self):
        self.match_patterns()
        for orderline in self.orderlines_list:
            orderline_parts = orderline[2:].split("\\n")

            self.orderlinenr = orderline_parts[0]
            self.item = orderline_parts[1]
            self.delivery_date = orderline_parts[2]
            self.qty = int(orderline_parts[3])
            self.price_per_piece = float(orderline_parts[4].replace(",", "."))
            values = [
                self.ordernumber,
                self.orderlinenr,
                self.item,
                self.delivery_date,
                self.qty,
                self.price_per_piece,
                self.delivery_terms,
                self.payment_terms,
            ]
            for idx, key in enumerate(self.purchaserow_extracted_dict):
                self.purchaserow_extracted_dict[key].append(values[idx])
        print(self.purchaserow_extracted_dict)
        self.purchaserow_extracted_df = pd.DataFrame(self.purchaserow_extracted_dict)
        return self.purchaserow_extracted_df

    def comparing_general_information(self, case, delivery_terms, payment_terms):
        process_dictonary = dict(
            zip(
                ["ordernumber", "delivery_terms", "payment_terms"],
                [self.ordernumber, delivery_terms, payment_terms],
            )
        )
        self.start_rpa_process(case)
        for process in process_dictonary:
            self.purchase_total_database_filtered_again = self.purchase_total_database_filtered[
                self.purchase_total_database_filtered[process] == process_dictonary[process]
            ]
            if len(self.purchase_total_database_filtered_again.index) == 0:
                for key in self.general_processes_dictonary[process][1:]:
                    self.time = self.time + timedelta(seconds=self.sleep_for_a_while())
                    self.process_mining_dict["case"].append(case)
                    self.process_mining_dict["timestamp"].append(self.time)
                    self.process_mining_dict["activity"].append(key)
            else:
                self.time = self.time + timedelta(seconds=self.sleep_for_a_while())
                self.process_mining_dict["case"].append(case)
                self.process_mining_dict["timestamp"].append(self.time)
                self.process_mining_dict["activity"].append(self.general_processes_dictonary[process][0])
        return self.process_mining_dict, self.time

    def comparing_orderlines(self, case, item, delivery_date, quantity, price_per_piece):
        if self.process_continue == True:
            process_dictonary = dict(
                zip(
                    ["item", "delivery_date", "quantity", "price_per_piece"],
                    [item, delivery_date, quantity, price_per_piece],
                )
            )
            for process in process_dictonary:
                self.purchase_total_database_filtered_again = self.purchase_total_database_filtered[
                    self.purchase_total_database_filtered[process] == process_dictonary[process]
                ]
                if len(self.purchase_total_database_filtered_again.index) == 0:
                    if process == "delivery_date":
                        run_dictionary = self.general_processes_dictonary[process][2:]
                    else:
                        run_dictionary = self.general_processes_dictonary[process][1:]
                    for key in run_dictionary:
                        self.time = self.time + timedelta(seconds=self.sleep_for_a_while())
                        self.process_mining_dict["case"].append(case)
                        self.process_mining_dict["timestamp"].append(self.time)
                        self.process_mining_dict["activity"].append(key)
                else:
                    if process == "delivery_date":
                        run_dictionary = self.general_processes_dictonary[process][0:2]
                    else:
                        run_dictionary = [self.general_processes_dictonary[process][0]]
                    for key in run_dictionary:
                        self.time = self.time + timedelta(seconds=self.sleep_for_a_while())
                        self.process_mining_dict["case"].append(case)
                        self.process_mining_dict["timestamp"].append(self.time)
                        self.process_mining_dict["activity"].append(key)
        return self.process_mining_dict

    #! Supporting functions
    def determine_to_continue_process(self):
        return self.process_continue

    def return_process_mining_dict(self):
        return self.process_mining_dict

    def sleep_for_a_while(self):
        self.sleeping = randint(30, 60)
        return self.sleeping

    def match_patterns(self):
        for orderline in range(1, 10):
            pattern = rf"((\D[n]{orderline}[.]\D[n]123)(.*?)(\D[n]€\s{{1,50}}\D[n][0-9,.]{{4,9}}\D[n]€))"
            matches = re.search(pattern, self.pymupdf_text)
            if matches:
                self.orderlines_raw = matches.group()
                self.orderlines_list.append(self.orderlines_raw)
        return self.orderlines_list

    def start_rpa_process(self, case):
        self.time = fake.date_time_between(
            start_date=datetime(2024, 2, 29, 7, 0, 0),
            end_date=datetime(2024, 3, 5, 17, 0, 0),
        )
        for idx, key in enumerate(self.general_processes_dictonary["process_start"]):
            if idx > 0:
                self.time = self.time + timedelta(seconds=self.sleep_for_a_while())
            self.process_mining_dict["case"].append(case)
            self.process_mining_dict["timestamp"].append(self.time)
            self.process_mining_dict["activity"].append(key)
        return self.process_mining_dict, self.time


if __name__ == "__main__":
    final_pm_df = pd.DataFrame()
    # pdfs = [f"{file_path}/fake_confirmations/pdfs/Confirmation_2410103_92408723.pdf"]
    pdfs = glob.glob(f"{file_path}/fake_confirmations/pdfs/*.pdf", recursive=True)
    for pdf in pdfs:
        reader = PurchaseOrderReader(pdf)
        reader.import_and_merge_snappies()
        reader.pymupdf_data_extractor()
        reader.extract_ordernumber_and_comparing_ordernumber_with_database()
        process_continue = reader.determine_to_continue_process()
        if process_continue == True:
            reader.extract_general_information()
            df = reader.transforming_orderlines_raw()
            print(df)
            for idx, row in df.iterrows():
                case = row["ordernumber"] + "-" + row["orderline"]
                ordernumber = row["ordernumber"]
                orderline = row["orderline"]
                item = row["item"]
                delivery_date = row["delivery_date"]
                quantity = row["quantity"]
                price_per_piece = row["price_per_piece"]
                delivery_terms = row["delivery_terms"]
                payment_terms = row["payment_terms"]
                reader.comparing_general_information(case, delivery_terms, payment_terms)
                reader.comparing_orderlines(case, item, delivery_date, quantity, price_per_piece)
        temp_df = pd.DataFrame(reader.return_process_mining_dict()).drop_duplicates()
        # print(temp_df)
        final_pm_df = pd.concat([final_pm_df, temp_df]).reset_index(drop=True)
final_pm_df.to_parquet(
    "fake_confirmations\snappy\process_mining_order_confirmation.snappy.parquet",
    compression="snappy",
    engine="pyarrow",
)
