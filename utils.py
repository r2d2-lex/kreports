#!/usr/bin/env python3
from html.parser import HTMLParser
import requests
import sys


def main_test():
    requests.packages.urllib3.disable_warnings()
    if len(sys.argv) > 1:
        for url in sys.argv[1:]:
            print(get_url_title(url))
    else:
        print('Enter url...  Example: shell# {} http://test.com/'.format(sys.argv[0]))


class MyHTMLParser(HTMLParser):
    def handle_endtag(self, tag):
        if tag == 'title':
            raise StopIteration()

    def handle_data(self, data):
        self.title = data


def get_url_title(url):
    http_proxy = "http://192.168.1.1:8080"
    https_proxy = "https://192.168.1.1:8080"
    ftp_proxy = "ftp://192.168.1.1:8080"
    proxy_dict = {
        "http": http_proxy,
        "https": https_proxy,
        "ftp": ftp_proxy
    }
    proxy_dict = None
    try:
        if proxy_dict:
            response = requests.get(url, stream=True, proxies=proxy_dict, verify=False, timeout=1)
        else:
            response = requests.get(url, stream=True, verify=False, timeout=1)
        print('Url: ', url, ' Response', response.status_code)

        # запрашиваем ровно 2048 байт, для чтения тега head этого должно хватать, или можно еще увеличить
        data = next(response.iter_content(2048))
        parser = MyHTMLParser()
        if response.encoding is None:
            response.close()
            return None
        rcv_data = data.decode(response.encoding)
        response.close()
    except requests.exceptions.SSLError:
        return None
    except requests.exceptions.ConnectionError:
        return None
    except StopIteration:
        return None

    try:
        parser.feed(rcv_data)
    except StopIteration:
        return parser.title
    return None


if __name__ == "__main__":
    main_test()
