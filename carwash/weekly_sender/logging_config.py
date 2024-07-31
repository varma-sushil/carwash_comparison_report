import logging
import logging.config
import os

from datetime import datetime

def setup_logging():
    log_folder = 'logs'
    if not os.path.exists(log_folder):
        os.makedirs(log_folder)
        
    # Create a filename with the current date and time
    log_filename = datetime.now().strftime("%Y-%m-%d_%H-%M-%S.log")

    logging_config = {
        'version': 1,
        'disable_existing_loggers': False,
        'formatters': {
            'standard': {
                'format': '%(asctime)s [%(levelname)s] %(name)s: %(message)s'
            },
        },
        'handlers': {
            'console': {
                'class': 'logging.StreamHandler',
                'formatter': 'standard',
                'level': 'INFO',
            },
            'file': {
                'class': 'logging.FileHandler',
                'filename': os.path.join(log_folder, log_filename),
                'formatter': 'standard',
                'level': 'DEBUG',
            },
        },
        'loggers': {
            '': {  # root logger
                'handlers': ['console', 'file'],
                'level': 'INFO',
                'propagate': True
            },
        }
    }

    logging.config.dictConfig(logging_config)