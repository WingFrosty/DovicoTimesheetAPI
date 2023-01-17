"""
Name        : utils_retry.py
Version     : 1.0
Author      : WingFrost
Description : Retry utilities

Date        Author                  Description
==========  ======================  ====================================================================================
16-01-2023  WingFrost               Initial Creation
"""
import time
import logging

logger = logging.getLogger(__name__)

def retry(max_tries, wait_interval_seconds):
    """
    Name: retry
    Purpose: Retry decoration for functions/methods
    Arguments:
        max_tries (int)
        wait_interval_seconds (int)
    """
    def decorator(func):
        def new_function(*args, **kwargs):
            tries = 0
            while (tries < max_tries):
                try:
                    tries += 1
                    return func(*args, **kwargs)  # end cycle if successful
                except:
                    if (tries == max_tries):  # raise exception at max_tries
                        raise
                    else:
                        logger.error("An exception was caught! Attempt nr. %s . Retrying again in %s seconds.",
                                       str(tries), str(wait_interval_seconds), exc_info=True)
                        time.sleep(wait_interval_seconds)
                        continue
        return new_function
    return decorator
