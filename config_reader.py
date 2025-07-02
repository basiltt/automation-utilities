from os import path
from typing import Any

import yaml

from app.models.config import (
    OcaNgqConfig,
    OcaConfig,
    NgqConfig,
)


class ConfigDataBlankError(Exception):
    """
    Exception raised when config data is blank
    """


class ConfigKeyMissingError(Exception):
    """
    Exception raised when config key is missing error
    """


# Defining config file
CONFIG_FILE = "config.yaml"
if not path.exists(CONFIG_FILE):
    raise FileNotFoundError(
        "config.yaml file not found in " "the base project directory"
    )


def read_config(
    tag: str,
    write_to_console: bool = False,
    raise_error_on_not_found: bool = False,
    validate_existence: bool = False,
    if_none_return: Any = None,
) -> Any:
    """
    Reads the config file and returns the value of the tag
    :param tag: Tag to be read as string
    :param write_to_console: If true, writes the value to console
    :param raise_error_on_not_found: If true, raises an error if the tag is not found
    :param validate_existence: If true, raises an error if the file is not found (Used for checking the file/folder -
        existence) as boolean
    :param if_none_return:  If the tag is not found, returns this value as default
    :return: Value of the tag
    """
    with open(CONFIG_FILE, "r") as file:
        config = yaml.load(file, Loader=yaml.FullLoader)
        if write_to_console:
            print(config[tag])
        if tag not in config.keys():
            raise ConfigKeyMissingError(f'"{tag}" not found in config file')
        if raise_error_on_not_found:
            if config[tag] is None:
                raise ConfigDataBlankError(f'"{tag}" is empty in config file')
        if validate_existence:
            if not path.exists(config[tag]):
                raise FileNotFoundError(f'"{config[tag]}" not found')
        if if_none_return:
            if config[tag] is None:
                return if_none_return
        return config[tag]


def get_config_data():
    """
    Get the config data from the config file
    :return: config object as DeploymentConfigModel
    """
    oca_xpath = read_config("OCA_XPATH")
    ngq_xpath = read_config("NGQ_XPATH")

    config_object = OcaNgqConfig(
        oca_config=OcaConfig(**oca_xpath),
        ngq_config=NgqConfig(**ngq_xpath),
    )
    return config_object
