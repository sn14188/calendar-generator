import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, timedelta

def load_and_clean_schedule(input_path):
    df = pd.read_excel(input_path)
    df.columns = df.iloc[1]
    df = df.iloc[2:, 1:]

    df = df[["Section", "Meeting Patterns"]]
    df["Section"] = df["Section"].str.split(" - ").str[0]
    df = df.rename(columns={"Meeting Patterns": "Meeting_Patterns"})

    df = (
        df.assign(
            Meeting_Patterns=df["Meeting_Patterns"]
            .str.split("\n")
            .apply(lambda x: [item for item in x if item.strip()])
        )
        .explode("Meeting_Patterns")
    )

    return df

def parse_schedule_details(df):
    df[["Period", "Days", "Time", "Location"]] = df["Meeting_Patterns"].str.split("|", expand=True)
    df[["Start Date", "End Date"]] = df["Period"].str.split(" - ", expand=True)
    df[["Start Time", "End Time"]] = df["Time"].str.split("-", expand=True)

    df["Start Time"] = df["Start Time"].apply(convert_time_format)
    df["End Time"] = df["End Time"].apply(convert_time_format)
    df["Days"] = df["Days"].str.strip()

    return df.drop(columns=["Meeting_Patterns", "Period", "Time"])

def convert_time_format(time_string):
    formatted_time = time_string.replace("a.m.", "AM").replace("p.m.", "PM").strip()
    return datetime.strptime(formatted_time, "%I:%M %p").strftime("%H:%M")

def create_calendar_events(df):
    calendar = Calendar()
    day_map = {
        "Mon": "MO",
        "Tue": "TU",
        "Wed": "WE",
        "Thu": "TH",
        "Fri": "FR",
        "Sat": "SA",
        "Sun": "SU",
    }

    for _, row in df.iterrows():
        start_date = datetime.strptime(row["Start Date"].strip(), "%Y-%m-%d")
        end_date = datetime.strptime(row["End Date"].strip(), "%Y-%m-%d") + timedelta(days=1)
        start_time = row["Start Time"]
        end_time = row["End Time"]
        location = row["Location"].strip()
        section = row["Section"].strip()
        days = row["Days"].strip()
        rrule_days = ",".join([day_map[day] for day in days.split()])
        start_datetime = datetime.combine(start_date, datetime.strptime(start_time, "%H:%M").time())
        end_datetime = datetime.combine(start_date, datetime.strptime(end_time, "%H:%M").time())

        event = Event()
        event.add("summary", section)
        event.add("dtstart", start_datetime)
        event.add("dtend", end_datetime)
        event.add("location", location)
        event.add("rrule", {"freq": "weekly", "byday": rrule_days.split(","), "until": end_date})

        calendar.add_component(event)
    return calendar

def save_calendar_to_file(calendar, output_path):
    output_file = f"{output_path}/courses.ics"
    with open(output_file, "wb") as f:
        f.write(calendar.to_ical())
    print(f"ics file saved at: {output_file}")

if __name__ == "__main__":
    input_path = "input/courses.xlsx"
    output_path = "output"

    df = load_and_clean_schedule(input_path)
    df = parse_schedule_details(df)
    calendar = create_calendar_events(df)
    save_calendar_to_file(calendar, output_path)
