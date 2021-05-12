from configparser import ConfigParser

__all__ = ["Create_Settings_File"]


def Create_Settings_File(gmail='', password=''):
    config = ConfigParser()

    config['settings'] = {
        "gmail": gmail,
        "password": password
    }

    with open('Account.ini', 'w') as fp:
        config.write(fp)
