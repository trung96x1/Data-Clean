import logging
from datetime import datetime
import os
from Define import *

class SingletonMeta(type):
    _instances = {}

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super().__call__(*args, **kwargs)
        return cls._instances[cls]

class Logger(metaclass=SingletonMeta):
    def __init__(self, name: str = DEFAULT_LOG_NAME, logLevel: int = logging.INFO):
        currentTime = datetime.now().strftime('%Y%m%d%H%M%S')
        logFoler = os.path.join(os.getcwd(), "Log")
        if not os.path.exists(logFoler):
            os.makedirs(logFoler)
        logPath = os.path.join(logFoler, f"{currentTime}.log")

        # This function config root logger, all 3rd party libraries will use this logger
        # logging.basicConfig(
        #     level=logLevel,  # Minimum level to capture
        #     format='[%(asctime)s] [%(levelname)s] %(name)s: %(message)s',
        #     datefmt='%Y-%m-%d %H:%M:%S',
        #     handlers=[
        #         logging.FileHandler(logPath),       # Log to file
        #         logging.StreamHandler()               # Log to console
        #     ]
        # )

        self.logger = logging.getLogger(name)
        self.logger.setLevel(logLevel)
        fileHandler = logging.FileHandler(logPath)
        # Even if the logger processes messages at the DEBUG level, 
        # This handler will only write messages to the file if they meet the level set here (DEBUG or higher in this case)
        fileHandler.setLevel(logLevel)
        fileHandler.setFormatter(logging.Formatter('[%(asctime)s] [%(levelname)s] %(name)s: %(message)s', datefmt='%Y-%m-%d %H:%M:%S'))
        self.logger.addHandler(fileHandler)                 # Add file handler
        self.logger.addHandler(logging.StreamHandler())     # Add console handler
    
    def get_logger(self):
        return self.logger

    def logd(self, message: str):
        self.logger.debug(message)

    def logi(self, message: str):
        self.logger.info(message)

    def logw(self, message: str):
        self.logger.warning(message)

    def loge(self, message: str):
        self.logger.error(message)