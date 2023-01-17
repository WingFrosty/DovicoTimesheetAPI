"""
Name        : utils_logger.py
Version     : 1.0
Author      : WingFrost
Description : Logger utilities

Date        Author                  Description
==========  ======================  ====================================================================================
16-01-2023  WingFrost               Initial Creation
"""
import sys
import os
import logging

default_log_level = logging.INFO
default_log_format = "[%(asctime)s] [%(name)s] [%(levelname)s] [PID=%(process)s] [THREAD=%(threadName)s] %(message)s"

def init_file_logger(log_file_path, log_level=default_log_level, log_format=default_log_format):
    """
    Name: init_file_logger
    Purpose: Initialize logger
    Arguments:
        log_file_path (str)
        log_level (int)
        log_format (str)
    """
    log_dir_path = os.path.dirname(log_file_path)

    #If directory does not exist, create it
    if not(os.path.exists(log_dir_path)):
        os.makedirs(log_dir_path)

    #Start logger
    logging.basicConfig(
        level=log_level,
        format=log_format,
        handlers=[
            logging.FileHandler(log_file_path, mode="a"),
            logging.StreamHandler(sys.stdout)
        ]
    )