import requests
from lxml import html as lxml_html
import browser_cookie3
import re
import pickle
import xlsxwriter
import random
import time

global_constants = None


def get_constants():
    """Return a dictionary containing all input constraints for scraping.
    All attributes except URL can be changed as per requirement"""

    dict_constants = dict()
    dict_constants['LOGIN_URL'] = "https://yocket.in/account/login"
    dict_constants['PAST_RESULTS_URL'] = "https://yocket.in/recent-admits-rejects?page="
    dict_constants['ALL_RESULTS_URL'] = "https://yocket.in/profiles/find/matching-admits-and-rejects?page="
    dict_constants['HOME_PAGE'] = 'https://yocket.in/'
    dict_constants['NUMBER_PAGE_TO_SCRAPE_FIRST'] = 1
    dict_constants['NUMBER_PAGE_TO_SCRAPE_LAST'] = 2
    dict_constants['MINIMUM_GPA'] = 7.5
    dict_constants['MINIMUM_GRE'] = 320
    dict_constants['MINIMUM_TOEFL'] = 100

    return dict_constants


def get_gpa(current_bucket_gpa):
    """Return a float value.
    Computed grade is obtained by getting numeric/floating
    part of string and then converting it into GPA if it is a
    percentage by a factor of 10."""

    computed_grade = re.findall(r"\d+[.]?\d*", current_bucket_gpa)
    if len(computed_grade) > 0:
        computed_grade = computed_grade[0]
        if float(computed_grade) > 10:
            computed_grade = float(computed_grade[0]) / 10.0
        return round(float(computed_grade), 2)
    return 0.0


def get_gre_or_toefl(current_bucket_marks):
    """Return a int value.
    Computed marks obtained by getting integer part of string
    Marks for IELTS will be converted to zero or be filtered at later
    stage."""

    computed_marks = current_bucket_marks.replace("\n", "").strip()
    try:
        int(computed_marks)
        return computed_marks
    except ValueError:
        return 0


def get_workex_months(current_bucket_workex):
    """Return a int value.
    Computed workex is obtained by getting numeric
    part of string which is equivalent to the number of months."""

    computed_workex = re.findall(r"\d+", current_bucket_workex)
    if len(computed_workex) > 0:
        return computed_workex[0]
    return 0


def filter_criteria_met(current_gre, current_gpa, current_toefl):
    """Return a boolean value.
    If either of minimum constraints not met then False returned."""

    if int(current_gre) < global_constants['MINIMUM_GRE']:
        return False
    if float(current_gpa) < global_constants['MINIMUM_GPA']:
        return False
    if int(current_toefl) < global_constants['MINIMUM_TOEFL']:
        return False
    return True

def split_bucket_university_course(current_bucket_university_course):
    """Return a 2 tuple list(university, course).
    Split performed on keywords which can be course starting names."""

    course_separator_delimiter = ['computer', 'artificial', 'cyber', 'network', 'data']

    for delimiter in course_separator_delimiter:
        separated_list = current_bucket_university_course.split(delimiter,1)
        if len(separated_list) == 2:
            return separated_list[0], delimiter+separated_list[1]

    return None, None


def export_to_file(final_data_fetch):
    """Export decision data to local files.

    First file created is excel file in readable format
    Second file is binary file which can be used for analytics."""

    # Column names for data
    header_fields = ['Course', 'University', 'GPA', 'GRE', 'TOEFL', 'Work Experience', 'UG Course', 'UG College','Admit Status']
    with xlsxwriter.Workbook('yocket_data.xlsx') as workbook:
        worksheet = workbook.add_worksheet()

        # Write Header Fields
        worksheet.write_row(0, 0, header_fields)
        # Write data fields
        for row_num, data in enumerate(final_data_fetch):
            worksheet.write_row(row_num+1, 0, data)

    # Store as binary data
    with open('yocket_data.data', 'wb') as f:
        pickle.dump(final_data_fetch, f)


def perform_scraping(current_session):
    """Trigger relevant HTTP calls to get requisite data
    and perform actual scraping"""

    # List Array storing all relevant decision information
    final_data_fetch = []
    pagination_index = global_constants['NUMBER_PAGE_TO_SCRAPE_FIRST']
    while pagination_index < global_constants['NUMBER_PAGE_TO_SCRAPE_LAST']:
        print("Page:", pagination_index, " Collected records:", len(final_data_fetch))

        # Get relevant admit-reject page based on pagination value
        result = current_session.get(global_constants['ALL_RESULTS_URL'] + str(pagination_index),
                                     headers=dict(referer=global_constants['ALL_RESULTS_URL']))
        tree = lxml_html.fromstring(result.content)

        # Get Nodes containing individual decisions for each page(approx 20 per page)
        decision_buckets = tree.xpath('//*[@class="row"]/div[@class="col-sm-6"]/div[@class="panel panel-warning"]/div[@class="panel-body"]')

        # If decision buckets are empty, captcha page has been encountered
        if len(decision_buckets) == 0:
            print("Captcha Time")
            time.sleep(120)
            continue

        for individual_decision_bucket in decision_buckets:

            current_admit_status = ((individual_decision_bucket.xpath('./div[1]/div[2]/label'))[0]).text.strip()

            # Fetch results only if ADMIT or REJECT
            if current_admit_status.lower() == 'admit' or current_admit_status.lower() == 'reject':

                # Get relevant information from html page returned in response
                current_bucket_university_course = ((individual_decision_bucket.xpath('./div[1]/div[1]/h4/small'))[0]).text.replace("\n","").strip()
                current_gre = get_gre_or_toefl(((((individual_decision_bucket.xpath('./div[2]/div[1]'))[0]).getchildren())[1]).tail)
                current_toefl = get_gre_or_toefl(((((individual_decision_bucket.xpath('./div[2]/div[2]'))[0]).getchildren())[1]).tail)
                current_gpa = get_gpa(((((individual_decision_bucket.xpath('./div[2]/div[3]'))[0]).getchildren())[1]).tail)
                current_workex = get_workex_months(((((individual_decision_bucket.xpath('./div[2]/div[4]'))[0]).getchildren())[1]).tail)

                current_university, current_course = split_bucket_university_course(current_bucket_university_course.lower())
                # Append decision information to final bucket only if minimum criteria met
                if current_university is not None and filter_criteria_met(current_gre, current_gpa, current_toefl):

                    # Get UG College from profile of user
                    profile_page_path = ((individual_decision_bucket.xpath('./div[1]/div[1]/h4/a'))[0]).attrib['href']
                    profile_result = current_session.get(global_constants['HOME_PAGE'] + profile_page_path,
                                                         headers=dict(referer=global_constants['PAST_RESULTS_URL']))
                    profile_tree = lxml_html.fromstring(profile_result.content)
                    ug_details_bucket = (profile_tree.xpath('//div[@class="col-sm-12 card"][1]'))
                    if len(ug_details_bucket) >= 1:
                        ug_details_bucket = ug_details_bucket[0]
                        current_ug_course = ((ug_details_bucket.xpath('./div[1]/div[7]/p[1]/b[1]'))[0]).text.replace("\n", "").strip()
                        current_ug_college = ((ug_details_bucket.xpath('./div[1]/div[7]/p[2]'))[0]).text.replace("\n", "").strip()

                        final_data_fetch.append([current_course, current_university, current_gpa, current_gre, current_toefl,
                                                 current_workex, current_ug_course, current_ug_college, current_admit_status])

        # Add sleep time to allow for web scraping in undetected manner
        sleep_delay = random.choice([0, 1, 2, 3])
        time.sleep(sleep_delay)
        pagination_index += 1

    # Export final_data to excel sheet
    export_to_file(final_data_fetch)


def main():
    # Get cookie from Chrome Browser
    cookiejar = browser_cookie3.chrome()

    # Load all scraping constants
    global global_constants
    global_constants = get_constants()

    # Start session
    current_session = requests.session()

    # Get login page and set cookie of current session as the browser session to bypass authentication
    current_session.get(global_constants['LOGIN_URL'], cookies=cookiejar, headers=dict([('referer', global_constants['HOME_PAGE'])]))

    # Perform actual scraping
    perform_scraping(current_session)


if __name__ == "__main__":
    main()
