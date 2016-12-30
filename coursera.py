import json
import sys
from random import sample

import requests
from bs4 import BeautifulSoup
from lxml import etree
from openpyxl import styles, Workbook


def get_courses_list(number_of_links):
    courses_xml = 'https://www.coursera.org/sitemap~www~courses.xml'
    request = requests.get(courses_xml)
    root = etree.fromstring(request.content)
    links = [loc.text for loc in root.iter('{*}loc')]
    random_links = sample(links, number_of_links)
    return random_links


def get_course_info(course_link):
    print("Processing: {}".format(course_link))

    course_data = requests.get(course_link).content
    soup = BeautifulSoup(course_data, 'html.parser')

    course_dict = {}

    course_dict['url'] = course_link

    course_title = soup.find("div", {"class": "title"}).text
    course_dict['title'] = course_title

    course_lang = soup.find("div", {"class": "language-info"}).text
    course_dict['lang'] = course_lang

    rating_tag = soup.find("div", {"class": "ratings-text bt3-visible-xs"})
    course_rating = 'no data'
    if rating_tag:
        course_rating = float(rating_tag.text.split(' ')[0])
    course_dict['rating'] = course_rating

    course_json_tag = soup.find('script', {'type': 'application/ld+json'})
    course_start_date = 'no data'
    if course_json_tag:
        course_date = json.loads(course_json_tag.text)['hasCourseInstance'][0]
        if 'startDate' in course_date:
            course_start_date = course_date['startDate']
    course_dict['start_date'] = course_start_date

    course_week_tag = soup.findAll('div', {'class': 'week-heading'})
    course_duration_weeks = 'no data'
    if course_week_tag:
        course_duration_weeks = len(course_week_tag)
    course_dict['duration_weeks'] = course_duration_weeks

    return course_dict


def get_courses_info(courses_links):
    return [get_course_info(course_link) for course_link in courses_links]


def output_courses_info_to_xlsx(filepath, courses_info):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = 'Courses'

    header = ['Course Title', 'Start date', 'Duration (weeks)', 'Language', 'Rating', 'URL']

    worksheet.append(header)

    for course_info in courses_info:
        worksheet.append(
            [course_info['title'],
             course_info['start_date'],
             course_info['duration_weeks'],
             course_info['lang'],
             course_info['rating'],
             course_info['url']]
        )

    workbook.save(filepath)

    print('Done! Data saved in {}!'.format(filepath))



if __name__ == '__main__':
    number_of_links = 1
    courses_links = get_courses_list(number_of_links)
    courses_info = get_courses_info(courses_links)
    filepath = sys.argv[1]
    output_courses_info_to_xlsx(filepath, courses_info)
