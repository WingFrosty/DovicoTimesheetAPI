import sys
import os
import logging

log_format = "[%(asctime)s] [%(name)s] [%(levelname)s] [PID=%(process)s] [THREAD=%(threadName)s] %(message)s"

def init_file_logger(log_file_path, logging_level):
    log_dir_path = os.path.dirname(log_file_path)

    #If directory does not exist, create it
    if not(os.path.exists(log_dir_path)):
        os.makedirs(log_dir_path)

    #Start logger
    logging.basicConfig(
        level=logging_level,
        format=log_format,
        handlers=[
            logging.FileHandler(log_file_path, mode="a"),
            logging.StreamHandler(sys.stdout)
        ]
    )