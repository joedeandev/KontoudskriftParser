from csv import writer
from datetime import datetime, timedelta
from decimal import Decimal
from os import listdir
from os.path import abspath, join
from re import IGNORECASE, match
from sys import argv
from typing import Union

from bs4 import BeautifulSoup
from fitz import Document

ENTRY_DATE_LEFT_POSITION = 57
VALUE_DATE_LEFT_POSITION = 98
CREDITED_MIN_LEFT_POSITION = 390
BALANCE_MIN_LEFT_POSITION = 490


class BankItemLine:
    def __init__(
        self,
        entry_time: datetime = None,
        value_time: datetime = None,
        description: list[str] = None,
        credited: Decimal = None,
        balance: Decimal = None,
    ):
        self.entry_time: datetime = entry_time
        self.value_time: datetime = value_time
        self.description: list[str] = [] if description is None else description
        self.credited: Decimal = credited
        self.balance: Decimal = balance

    def is_complete(self) -> bool:
        if self.entry_time is None:
            return False
        if self.value_time is None:
            return False
        if self.credited is None:
            return False
        if self.balance is None:
            return False
        if len(self.description) < 1:
            return False
        return True

    def __repr__(self):
        return (
            f"{self.entry_time=} {self.value_time=} "
            f"{self.description=} "
            f"{self.credited=} {self.balance=}"
        )

    def __str__(self):
        return "\t".join([str(i) for i in self.as_tuple()])

    def as_tuple(self) -> tuple[str, str, str, str, str]:
        return (
            self.entry_time.strftime("%Y/%m/%d"),
            self.entry_time.strftime("%Y/%m/%d"),
            " ".join(self.description),
            str(None) if self.credited is None else str.format("{:.2f}", self.credited),
            str(None) if self.balance is None else str.format("{:.2f}", self.balance),
        )

    def __gt__(self, other: "BankItemLine"):
        return self.entry_time.__gt__(other.entry_time)

    def __lt__(self, other):
        return self.entry_time.__lt__(other.entry_time)


def bank_currency_format_to_decimal(text: str) -> Decimal:
    d = Decimal("".join([i for i in text if i in [c for c in "1234567890."]]))
    if text.endswith("+"):
        d = d * 1
    elif text.endswith("-"):
        d = d * -1
    else:
        raise Exception(f"Unknown end character on currency {text}")
    return d


def parse_doc(filepath) -> list[BankItemLine]:
    doc = Document(filepath)
    items = []
    for page_number in range(doc.page_count):
        page = doc.load_page(page_number).get_text("html")
        items += parse_page(page)
    return sorted(items)


def parse_page(page: str) -> list[BankItemLine]:
    html = BeautifulSoup(page, features="html.parser")

    items: list[BankItemLine] = []

    record: Union[BankItemLine, None] = None
    is_recording = False
    time_range_start: Union[datetime, None] = None
    time_range_end: Union[datetime, None] = None

    # iterate through elements in the page - luckily, page data is at least somewhat
    # sequentially ordered in elements
    for i in html.find_all("p"):
        # filter out most of the non-relevant elements,
        # leaving us with a soup of the relevant elements
        try:
            s = i.find("span")
            if not s:
                raise Exception
            text: str = s.text
            span_style: str = s["style"]
            par_style: str = i["style"]

            if "font-size:9pt" not in span_style:
                raise Exception
        except:
            continue

        # "Balance as at 30. 11. 2016" indicates that there is no more useful info
        if time_range_end is not None and text.startswith(
            f'Balance as at {time_range_end.strftime("%d. %m. %Y")}'
        ):
            is_recording = False
            continue

        # "Period this statement relates to: 01.09.2016 to 30.11.2016" indicates the
        # time that dates are relative to and the start of the section of useful info
        if not is_recording and text.startswith("Period this statement relates to"):
            dates = text.split(": ")[1].split(" to ")
            time_range_start = datetime.strptime(dates[0], "%d.%m.%Y")
            time_range_end = datetime.strptime(dates[1], "%d.%m.%Y")
            is_recording = True
            continue

        # ignore non-recorded, non-recording-control elements
        if not is_recording:
            continue

        # determine the element "type" by its left style position
        left_style = match(r"^.*left:(\d+)pt", par_style)
        if not left_style:
            raise Exception(f"No left style in paragraph style ({par_style})")
        left_amount = int(left_style.group(1))

        if left_amount in [ENTRY_DATE_LEFT_POSITION, VALUE_DATE_LEFT_POSITION]:
            is_entry = left_amount == ENTRY_DATE_LEFT_POSITION
            time_range_end: datetime = (
                time_range_start if time_range_end is None else time_range_end
            )
            # non-date items in same position
            if not match(r"^\d\d\.\d\d$", text):
                continue

            # the timestamp doesn't include the year, so we need to guess a bit
            # by comparing it to the start and end dates
            valid_stamp = None
            # "entry date" should always be within the range of the statement period...
            if is_entry:
                for year in [time_range_start.year, time_range_end.year]:
                    try:
                        timestamp = datetime.strptime(f"{text}.{year}", "%d.%m.%Y")
                    # ValueError thrown if given string is an invalid date (like 48.00)
                    except ValueError:
                        break
                    if time_range_start <= timestamp <= time_range_end:
                        valid_stamp = timestamp
                        break
            # but "value date" should be within a week of the entry
            else:
                for year in [record.entry_time.year, record.entry_time.year + 1]:
                    timestamp = datetime.strptime(f"{text}.{year}", "%d.%m.%Y")
                    if timestamp - record.entry_time < timedelta(days=7):
                        valid_stamp = timestamp
                        break
            if valid_stamp is None:
                raise Exception(
                    f"Timestamp {text} could not be assigned within range "
                    f"{time_range_start.strftime('%Y/%m/%d')} - "
                    f"{time_range_end.strftime('%Y/%m/%d')}"
                )
            # indicates the start of a new line item
            if is_entry:
                if record is None:
                    record = BankItemLine()
                else:
                    if record.is_complete():
                        items.append(record)
                        record = BankItemLine()
                    else:
                        raise Exception("Incomplete record found new entry timestamp")
                record.entry_time = valid_stamp
            else:
                if record is None or record.entry_time is None:
                    continue
                record.value_time = valid_stamp
            continue

        # only accept new items if the value entry timestamp has been added
        if record is None or record.value_time is None:
            continue
        # "balance" text
        elif left_amount > BALANCE_MIN_LEFT_POSITION:
            record.balance = bank_currency_format_to_decimal(text)
        # "credited" text
        elif left_amount > CREDITED_MIN_LEFT_POSITION:
            record.credited = bank_currency_format_to_decimal(text)
        # description text
        else:
            record.description.append(text)

    if record is None:
        pass
    elif record.is_complete():
        items.append(record)
    elif record.entry_time is not None:
        raise Exception("Page ended on incomplete record")

    return sorted(items)


if __name__ == "__main__":
    root_directory_path = argv[1] if len(argv) >= 2 else "."
    kontoudskrift_filename_regex = r"\d+ Kontoudskrift \d+\.pdf"
    kontoudskrift_items = []
    for kontoudskrift_filepath in [
        abspath(join(root_directory_path, f))
        for f in listdir(root_directory_path)
        if match(kontoudskrift_filename_regex, f, IGNORECASE)
    ]:
        kontoudskrift_items += parse_doc(kontoudskrift_filepath)
    kontoudskrift_items.sort()
    with open(
        join(root_directory_path, "kontoudskrift.csv"), "w", newline=""
    ) as output_file:
        csv_writer = writer(output_file)
        for item_line in kontoudskrift_items:
            csv_writer.writerow(list(item_line.as_tuple()))
