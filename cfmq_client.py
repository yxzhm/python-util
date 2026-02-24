import logging
import requests


host = 'https://vyrvk.wiremockapi.cloud'
destination = 'PBC.103002.OUT.0001'
logging.basicConfig(level=logging.DEBUG)


def _login_cfmq():
    url = f'{host}/login'
    header = {
        'CFMQ-Username': "SU1UU",
        'CFMQ-Password': "aW10c2VydmV"
    }

    response = requests.post(url=url, headers=header)
    response_json = response.json()
    return response_json['data']['CFMQ-Token']

if __name__ == '__main__':
    token = _login_cfmq()
    print(token)