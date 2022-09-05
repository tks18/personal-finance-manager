import appdirs, yaml
import os
import win32gui, win32con

APP_NAME = "Personal Finance Manager"
APP_AUTHOR = "tks18"


def get_config_dir():
    return appdirs.user_config_dir(APP_NAME, APP_AUTHOR)


def get_file_from_dialog(filter, file, extn, title):
    fileData = win32gui.GetOpenFileNameW(
        Flags=win32con.OFN_EXPLORER,
        File=file,
        DefExt=extn,
        Title=title,
        Filter=filter,
        FilterIndex=0,
    )
    return fileData[0]


def create_config(config_dir: str):
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)
    db_location = os.path.join(config_dir, "transactions.db")
    master_file_location = get_file_from_dialog(
        filter="Excel Files*.xlsx",
        file="Master",
        extn="xlsx",
        title="Select the Masters File",
    )
    data_to_store = {
        "DB": {"PATH": db_location},
        "MASTERS": {"PATH": master_file_location},
    }
    config = {}
    with open(os.path.join(config_dir, "app.yaml"), "x") as stream:
        try:
            yaml.dump(data_to_store, stream)
            config = data_to_store
        except:
            print("Error While Creating the Config File for Application")
    return config


def load_config(config_dir: str):
    config = {}
    with open(os.path.join(config_dir, "app.yaml"), "r") as stream:
        try:
            config = yaml.safe_load(stream=stream)
        except yaml.YAMLError as exc:
            create_config(config_dir=config_dir)
    return config


def load_or_create_config():
    config_dir = get_config_dir()
    config = {}
    if os.path.exists(os.path.join(config_dir, "app.yaml")):
        config = load_config(config_dir=config_dir)
    else:
        config = create_config(config_dir=config_dir)
    return config
