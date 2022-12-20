from unittest import TestCase, main
from main import Vacancy, DataSet, Report

vacancy_dct = {'name': 'Программист', 'salary_from': '10', 'salary_to': '20', 'salary_currency': 'RUR',
               'area_name': 'Екатеринбург', 'published_at': '2007-12-03T17:40:09+0300'}


class VacancyTest(TestCase):
    def test_vacancy_type(self):
        self.assertEqual(type(Vacancy(vacancy_dct)).__name__, 'Vacancy')

    def test_vacancy_name(self):
        self.assertEqual(Vacancy(vacancy_dct).name, 'Программист')

    def test_vacancy_salary_to(self):
        self.assertEqual(Vacancy(vacancy_dct).salary_to, 20)

    def test_vacancy_salary_from(self):
        self.assertEqual(Vacancy(vacancy_dct).salary_from, 10)

    def test_vacancy_salary_currency(self):
        self.assertEqual(Vacancy(vacancy_dct).salary_currency, 'RUR')

    def test_vacancy_average_salary(self):
        self.assertEqual(Vacancy(vacancy_dct).get_average_rub_salary(), 15.0)

    def test_vacancy_published_year(self):
        self.assertEqual(Vacancy(vacancy_dct).get_published_vacancy_year(), 2007)


class DatasetTest(TestCase):
    def test_dataset_type(self):
        self.assertEqual(type(DataSet('', '')).__name__, 'DataSet')

    def test_dataset_file_name(self):
        self.assertEqual(DataSet('vacancies_by_year.csv', 'Аналитик').file_name, 'vacancies_by_year.csv')

    def test_dataset_vacancy_name(self):
        self.assertEqual(DataSet('vacancies_by_year.csv', 'Аналитик').vacancy_name, 'Аналитик')


class ReportTest(TestCase):
    def test_report_type(self):
        self.assertEqual(type(Report('', {}, {}, {}, {}, {}, {})).__name__, 'Report')


if __name__ == '__main__':
    main()
