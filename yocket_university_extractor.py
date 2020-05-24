import requests
import browser_cookie3
import yocket_general_extractor as yge
import time
from lxml import html as lxml_html
import random
import pickle
import xlsxwriter
import re

global_constants = None


def get_constants():
    """Return a dictionary containing all input constraints for scraping.
    Courses can be added based on yocket URLs"""

    dict_constants = dict()
    dict_university_course_url = dict()
    dict_university_course_url['USC_General'] = 'https://yocket.in/applications-admits-rejects/168-university-of-southern-california/'
    dict_university_course_url['USC_DataScience'] = 'https://yocket.in/applications-admits-rejects/51720-university-of-southern-california/'
    dict_university_course_url['UCLA_General'] = 'https://yocket.in/applications-admits-rejects/53-university-of-california-los-angeles/'
    dict_university_course_url['UMass_General'] = 'https://yocket.in/applications-admits-rejects/232-university-of-massachusetts-amherst/'
    dict_university_course_url['Georgia_General'] = 'https://yocket.in/applications-admits-rejects/18-georgia-institute-of-technology/'
    dict_university_course_url['CMU_General'] = 'https://yocket.in/applications-admits-rejects/9-carnegie-mellon-university/'
    dict_university_course_url['CMU_ML'] = 'https://yocket.in/applications-admits-rejects/835-carnegie-mellon-university/'
    dict_university_course_url['CMU_MCDS'] = 'https://yocket.in/applications-admits-rejects/55949-carnegie-mellon-university/'
    dict_university_course_url['UCSD_General'] = 'https://yocket.in/applications-admits-rejects/219-university-of-california-san-diego/'
    dict_university_course_url['CalTech_GeneralPhD'] = 'https://yocket.in/applications-admits-rejects/226-california-institute-of-technology/'
    dict_university_course_url['UTA_General'] = 'https://yocket.in/applications-admits-rejects/46152-university-of-texas-austin/'
    dict_university_course_url['SBU_General'] = 'https://yocket.in/applications-admits-rejects/129-state-university-of-new-york-at-stony-brook/'
    dict_university_course_url['NYU_General'] = 'https://yocket.in/applications-admits-rejects/588-new-york-university/'
    dict_university_course_url['UIUC_General'] = 'https://yocket.in/applications-admits-rejects/81-university-of-illinois-at-urbana-champaign/'
    dict_university_course_url['Cornell_General'] = 'https://yocket.in/applications-admits-rejects/239-cornell-university/'
    dict_university_course_url['UMCP_General'] = 'https://yocket.in/applications-admits-rejects/31922-university-of-maryland-college-park/'

    dict_constants['LOGIN_URL'] = "https://yocket.in/account/login"
    dict_constants['HOME_PAGE'] = 'https://yocket.in/'
    dict_constants['admit_url_code'] = '2'
    dict_constants['reject_url_code'] = '3'
    dict_constants['pagination_suffix'] = '?page='
    dict_constants['course_url'] = dict_university_course_url
    dict_constants['MINIMUM_GPA'] = 7.5
    dict_constants['MINIMUM_GRE'] = 315
    dict_constants['MINIMUM_TOEFL'] = 100

    return dict_constants


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


def export_to_file(final_data_fetch, university_course):
    """Export decision data to local files corresponding to university and course.

    First file created is excel file in readable format
    Second file is binary file which can be used for analytics."""

    # Column names for data
    header_fields = ['Course', 'University', 'GPA', 'GRE Quant', 'GRE Verbal', 'TOEFL', 'Work Experience', 'UG Course', 'UG College', 'Admit Status',
                     'Papers', 'Profile']
    with xlsxwriter.Workbook('C:/Users/i349223/Downloads/YocketCode/ResultDocuments/' + university_course + '.xlsx') as workbook:
        worksheet = workbook.add_worksheet()

        # Write Header Fields
        worksheet.write_row(0, 0, header_fields)
        # Write data fields
        for row_num, data in enumerate(final_data_fetch):
            worksheet.write_row(row_num + 1, 0, data)

    # Store as binary data
    with open('C:/Users/i349223/Downloads/YocketCode/ResultDocuments/' + university_course + '.data', 'wb') as f:
        pickle.dump(final_data_fetch, f)


def extract_gre_partial_score(input_text):
    """Input is text containing numeric component containing GRE Quant or Verbal Score
    Output is the value if existing"""
    computed_partial_gre_score = re.findall(r"\d+", input_text)
    if computed_partial_gre_score is not None:
        return computed_partial_gre_score[0]
    return 0


def perform_scraping(current_session):
    """Trigger relevant HTTP calls to get requisite data
    and perform actual scraping"""

    for course_value in global_constants['course_url'].values():

        # List Array storing all relevant decision information for university anc course
        final_data_fetch = []

        # Get records for admit and reject
        for decision_code in ['admit_url_code', 'reject_url_code']:

            # Get all pages for decision
            pagination_index = 1
            while pagination_index < 300:
                print("Page:", course_value + global_constants[decision_code] + global_constants['pagination_suffix'] + str(pagination_index),
                      " Collected records:", len(final_data_fetch))
                # Get relevant admit-reject page based on pagination value
                result = current_session.get(course_value + global_constants[decision_code] + global_constants['pagination_suffix'] +
                                             str(pagination_index), headers=dict(referer=course_value))
                tree = lxml_html.fromstring(result.content)

                # Get Nodes containing individual decisions for each page(approx 20 per page)
                decision_buckets = tree.xpath('//*[@class="row"]/div[@class="col-sm-6"]/div[@class="panel panel-warning"]/div[@class="panel-body"]')

                # If decision buckets are empty, captcha page has been encountered or no limit has been reached
                if len(decision_buckets) == 0:
                    # No decision in further pages where status code 403
                    if result.status_code != 200:
                        break
                    # No decision in further pages if status code 200 and no results returned
                    no_results_page = (tree.xpath('//p[@class="lead"]/i'))
                    if len(no_results_page) > 0 and str(no_results_page[0].tail).strip().lower() == 'no matching profiles found!':
                        break
                    # Captcha Page
                    print("Captcha Time")
                    time.sleep(120)
                    continue

                for individual_decision_bucket in decision_buckets:

                    current_admit_status = ((individual_decision_bucket.xpath('./div[1]/div[2]/label'))[0]).text.strip()

                    # Fetch results only if ADMIT or REJECT
                    if current_admit_status.lower() == 'admit' or current_admit_status.lower() == 'reject':

                        # Get relevant information from html page returned in response
                        current_bucket_university_course = ((individual_decision_bucket.xpath('./div[1]/div[1]/h4/small'))[0]).text.replace("\n",
                                                                                                                                            "").strip()
                        current_gre = yge.get_gre_or_toefl(((((individual_decision_bucket.xpath('./div[2]/div[1]'))[0]).getchildren())[1]).tail)
                        current_toefl = yge.get_gre_or_toefl(((((individual_decision_bucket.xpath('./div[2]/div[2]'))[0]).getchildren())[1]).tail)
                        current_gpa = yge.get_gpa(((((individual_decision_bucket.xpath('./div[2]/div[3]'))[0]).getchildren())[1]).tail)
                        current_workex = yge.get_workex_months(((((individual_decision_bucket.xpath('./div[2]/div[4]'))[0]).getchildren())[1]).tail)

                        current_university, current_course = yge.split_bucket_university_course(current_bucket_university_course.lower())
                        # Append decision information to final bucket only if minimum criteria met
                        if current_university is not None and filter_criteria_met(current_gre, current_gpa, current_toefl):

                            # Add sleep time to allow for web scraping in undetected manner
                            sleep_delay = random.choice([1, 2, 3])
                            time.sleep(sleep_delay)

                            # Get UG College from profile of user
                            profile_page_path = ((individual_decision_bucket.xpath('./div[1]/div[1]/h4/a'))[0]).attrib['href']
                            profile_result = current_session.get(global_constants['HOME_PAGE'] + profile_page_path,
                                                                 headers=dict(referer=global_constants['HOME_PAGE']))
                            profile_tree = lxml_html.fromstring(profile_result.content)
                            ug_details_bucket = (profile_tree.xpath('//div[@class="col-sm-12 card"][1]'))

                            # Check if profile page exists
                            if len(ug_details_bucket) >= 1:
                                ug_details_bucket = ug_details_bucket[0]

                                # Get data if supplied
                                current_ug_course = ((ug_details_bucket.xpath('./div[1]/div[7]/p[1]/b[1]'))[0]).text
                                if current_ug_course is None:
                                    current_ug_course = ""
                                else:
                                    current_ug_course = current_ug_course.replace("\n", "").strip()

                                current_ug_college = ((ug_details_bucket.xpath('./div[1]/div[7]/p[2]'))[0]).text
                                if current_ug_college is None:
                                    current_ug_college = ""
                                else:
                                    current_ug_college = current_ug_college.replace("\n", "").strip()

                                current_papers = ((profile_tree.xpath('//div[@class="row text-center"]/div[4]/h4[1]/br'))[0]).tail
                                if current_papers is None:
                                    current_papers = ""
                                else:
                                    current_papers = current_papers.replace("\n", "").strip()

                                profile_gre_details_bucket = (profile_tree.xpath('//div[@id="yocket_app"]/div[@class="col-sm-6"]/div['
                                                                                 '@class="col-sm-12"]/div[@class="row text-center"]'))[0]
                                current_gre_quant = profile_gre_details_bucket.xpath('./div[1]/h4[1]/span[1]')
                                if len(current_gre_quant) > 0:
                                    current_gre_quant = extract_gre_partial_score(current_gre_quant[0].text)
                                    current_gre_verbal = extract_gre_partial_score(((profile_gre_details_bucket.xpath('./div[1]/h4[1]/span[1]/br[1]'))
                                                    [0]).tail)
                                    final_data_fetch.append([current_course, current_university, current_gpa, current_gre_quant, current_gre_verbal,
                                                         current_toefl, current_workex, current_ug_course, current_ug_college, current_admit_status,
                                                         current_papers, profile_page_path])

                pagination_index += 1

                # Add sleep time to allow for web scraping in undetected manner
                sleep_delay = random.choice([1, 2, 3])
                time.sleep(sleep_delay)

        # Export final_data to excel sheet
        export_to_file(final_data_fetch, str(current_university) + str(current_course))


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

    # Do actual scraping
    perform_scraping(current_session)


if __name__ == "__main__":
    main()
