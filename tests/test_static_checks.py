"""Быстрые проверки структуры без запуска БД. Запуск: python -m unittest discover -s tests"""
import ast
import unittest
from pathlib import Path
from jinja2 import Environment, FileSystemLoader

ROOT = Path(__file__).resolve().parents[1]

class ProjectStaticChecks(unittest.TestCase):
    def test_python_syntax(self):
        ast.parse((ROOT / 'app.py').read_text(encoding='utf-8'))

    def test_templates_compile(self):
        env = Environment(loader=FileSystemLoader(ROOT / 'templates'))
        for path in (ROOT / 'templates').glob('*.html'):
            env.get_template(path.name)

    def test_required_schema_tables_are_present(self):
        source = (ROOT / 'app.py').read_text(encoding='utf-8')
        for table in ('Departments','DepartmentsSections','Employees','EmployeeRoles','WorkTasks','WorkTaskExecutors','WorkTaskStatuses','WorkTaskHistory','WorkTaskComments','WorkTaskFiles'):
            self.assertIn("__tablename__='" + table + "'", source)

    def test_no_node_runtime_required(self):
        self.assertFalse((ROOT / 'package.json').exists())

if __name__ == '__main__':
    unittest.main()
