from typing import Callable
from datetime import datetime, timedelta


def execution_time(func: Callable) -> Callable:
    def wrapper(*args, **kwargs):
        initial_time: datetime = datetime.now()
        func(*args, **kwargs)
        final_time: datetime = datetime.now()
        time_elapsed: timedelta = final_time - initial_time
        print(f'Time passed {time_elapsed.total_seconds()} s.')

    return wrapper
