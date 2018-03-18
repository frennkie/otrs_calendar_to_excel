import calendar
import tablib
from datetime import datetime
import pymysql

try:
    from config import *
except ImportError:
    MYSQL_HOST = "localhost"
    MYSQL_USER = "otrs"
    MYSQL_PASS = "otrs"
    MYSQL_DB = "otrs"
    CALENDER_ID = 1
    APPOINTMENT_SYMBOL = "U"


def normalize_export(export):
    """take list of dict and convert to dict (key is the resource) with a list of dicts (start + end)
    in:
    [
        {"start_time": "2018-01-04 00:00:00", "end_time": "2018-01-06 00:00:00", "resource_id": "Person A"},
        {"start_time": "2018-01-30 00:00:00", "end_time": "2018-01-31 00:00:00", "resource_id": "Person A"},
        {"start_time": "2018-01-13 00:00:00", "end_time": "2018-01-14 00:00:00", "resource_id": "Person B"},
        {"start_time": "2018-01-04 00:00:00", "end_time": "2018-01-11 00:00:00", "resource_id": "Person B"}
    ]
    out:
    {'Person A': [{'start_time': datetime.date(2018, 1, 4), 'end_time': datetime.date(2018, 1, 6)},
                  {'start_time': datetime.date(2018, 1, 30), 'end_time': datetime.date(2018, 1, 31)}],
     'Person B': [{'start_time': datetime.date(2018, 1, 13), 'end_time': datetime.date(2018, 1, 14)},
                  {'start_time': datetime.date(2018, 1, 4), 'end_time': datetime.date(2018, 1, 11)}]}
    """
    normalized = {}
    for appointment in export:
        try:
            if isinstance(appointment["start_time"], datetime):
                normalized.setdefault(int(appointment["resource_id"]), []).append(
                    {
                        "start_time": appointment["start_time"].date(),
                        "end_time": appointment["end_time"].date()
                    }
                )
            else:
                normalized.setdefault(int(appointment["resource_id"]), []).append(
                    {
                        "start_time": datetime.strptime(appointment["start_time"], "%Y-%m-%d %H:%M:%S").date(),
                        "end_time": datetime.strptime(appointment["end_time"], "%Y-%m-%d %H:%M:%S").date()
                    }
                )
        except TypeError:
            # ignore appointments with an empty `resource_id`. int() will raise TypeError on None
            pass

    return normalized


def check_date_against_export_for_resource(date_obj, res_id, appointments):
    try:
        for appointment in appointments[res_id]:
            app_start = appointment["start_time"]
            app_end = appointment["end_time"]
            if app_start <= date_obj <= app_end:
                return True
    except KeyError:
        # if a resource has no appointments then it will not be in the dict -> that's ok
        pass

    return False


def get_agents(con):
    """return dict with `users`.`id` value as key
    e.g. {2: {'id': 2, 'login': 'fbar', 'last_name': 'foo', 'first_name': 'barr', 'mail': b'f@bar.com'}}
    """
    try:
        with con.cursor() as cur:
            # Create a new record
            sql = ("""SELECT users.id, users.login, users.last_name, users.first_name, 
                        CONCAT(users.last_name, ", ", users.first_name) as "display_name", u.preferences_value as 'mail'
                      FROM users
                      INNER JOIN user_preferences u ON users.id = u.user_id
                      WHERE preferences_key = 'UserEmail'
                      AND users.valid_id = 1
                      ORDER BY users.last_name""")

            cur.execute(sql)

        return {x['id']: x for x in cur.fetchall()}

    except Exception as err:
        raise("Error: {}".format(err))


def get_calendars(con):
    """return dict with `calendar`.`id` value as key
    e.g. {1: {'id': 1, 'name': 'Foobar'}, {'id': 2, 'name': 'Barfoo'}}
    """
    try:
        with con.cursor() as cur:
            # Create a new record
            sql = ("""SELECT id, name
                      FROM calendar
                      WHERE calendar.valid_id = 1
                      ORDER BY id ASC""")

            cur.execute(sql)

        return {x['id']: x for x in cur.fetchall()}

    except Exception as err:
        raise("Error: {}".format(err))


def get_calendar_appointments(con, cal_id):
    """return list of appointments
    e.g. {1: {'id': 1, 'name': 'Foobar'}, {'id': 2, 'name': 'Barfoo'}}
    """
    try:
        with con.cursor() as cur:
            # Create a new record
            sql = ("""SELECT id, title, start_time, end_time, all_day, team_id, resource_id, create_by, change_by
                      FROM calendar_appointment
                      WHERE calendar_id = %(calendar_id)s
                      ORDER BY id ASC""")

            cur.execute(sql, {"calendar_id": cal_id})

        # return {x['id']: x for x in cur.fetchall()}
        return cur.fetchall()

    except Exception as err:
        raise("Error: {}".format(err))


def main():
    # Connect to the database
    connection = pymysql.connect(host=MYSQL_HOST,
                                 user=MYSQL_USER,
                                 password=MYSQL_PASS,
                                 db=MYSQL_DB,
                                 charset='utf8mb4',
                                 cursorclass=pymysql.cursors.DictCursor)

    resources = ([(x[1]['id'], x[1]['display_name']) for x in get_agents(connection).items()])
    vacation_appointments = get_calendar_appointments(connection, CALENDER_ID)
    vacation_appointments_by_resource = normalize_export(vacation_appointments)
    # print(get_agents(connection))
    # print(resources)
    # print(get_calendars(connection))
    # print(vacation_appointments)
    # print(vacation_appointments_by_resource)

    headers = ["Date", "Day"]
    headers.extend([x[1] for x in resources])
    headers.append("Away (%)")

    data_cur = tablib.Dataset(headers=headers)
    data_last = tablib.Dataset(headers=headers)
    data_next = tablib.Dataset(headers=headers)

    c = calendar.TextCalendar(calendar.MONDAY)

    # create sheets for this year, last year and next year
    year = datetime.utcnow().year
    for data, year in ((data_cur, year), (data_last, year - 1), (data_next, year + 1)):
        # for month in range(1, 13):
        for month in range(1, 5):  # TODO (for dev only use Jan - April)
            data.append_separator("")
            data.append_separator("{} {}".format(calendar.month_name[month], year))
            for item in c.itermonthdates(year, month):
                if item.month == month:
                    date = item.strftime("%d.%m.%Y")
                    day = calendar.day_abbr[item.weekday()]
                    row = [date, day]

                    appointments_per_day = 0
                    for resource_id, resource_name in resources:
                        if check_date_against_export_for_resource(item, resource_id, vacation_appointments_by_resource):
                            row.append(APPOINTMENT_SYMBOL)
                            appointments_per_day += 1
                        else:
                            row.append("")

                    # percentage of resources that have an appointment
                    row.append("{0:.2f}".format(appointments_per_day / len(resources) * 100))

                    data.append(row)

    book = tablib.Databook((data_cur, data_last, data_next))

    with open('calendar.xls', 'wb') as f:
        # XLSX currently not supported https://github.com/kennethreitz/tablib/issues/324
        f.write(book.xls)


if __name__ == "__main__":
    main()
