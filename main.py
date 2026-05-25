from datetime import date
from http.server import HTTPServer, SimpleHTTPRequestHandler
from pathlib import Path

from jinja2 import Environment, FileSystemLoader

from read_wine2 import (
    URL,
    XLSX_PATH,
    download_xlsx,
    group_by_category,
    read_excel_rows,
)

BASE_DIR = Path(__file__).resolve().parent


def decline_year(num: int) -> str:
    if num % 100 in (11, 12, 13, 14):
        return "лет"
    
    last_digit = num % 10
    if last_digit == 1:
        return "год"
    elif last_digit in (2, 3, 4):
        return "года"
    else:
        return "лет"


class JinjaHTTPRequestHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/' or self.path == '/template.html':
            template_dir = str(BASE_DIR)
            env = Environment(loader=FileSystemLoader(template_dir), autoescape=True)
            template = env.get_template('template.html')

            current_year = date.today().year
            winery_age = current_year - 1920

            product_categories = self.load_product_categories()

            rendered = template.render(
                winery_age=winery_age,
                current_year=current_year,
                decline_year=decline_year,
                product_categories=product_categories,
            )

            self.send_response(200)
            self.send_header('Content-type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(rendered.encode('utf-8'))
        else:
            super().do_GET()

    def load_product_categories(self) -> dict:
        try:
            download_xlsx(URL, XLSX_PATH)
            wines = read_excel_rows(XLSX_PATH)
            return group_by_category(wines)
        except Exception as exc:
            print('Failed to load product categories:', exc)
            return {}


if __name__ == '__main__':
    server = HTTPServer(('0.0.0.0', 8000), JinjaHTTPRequestHandler)
    server.serve_forever()
