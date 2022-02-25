import requests
import pandas as pd
import numpy as np
import datatable as dt

from typing import List
from os import path
from lxml import html


def reduce_memory_usage(
        df: pd.DataFrame,
        verbose: bool = True
    ) -> pd.DataFrame:
    """
    Reduces memory usage.

    This function reduces memory usage by transforming the data types to a lesser cost in RAM,
    source: https://towardsdatascience.com/6-pandas-mistakes-that-silently-tell-you-are-a-rookie-b566a252e60d

    Parameters:
    - df: pd.DataFrame: A pandas dataframe.
    - verbose: bool: Prints the information of memory reduced. By default is True.

    Returns:
    - The dataframe optmized for memory utilization.
    """
    numerics = [
        "int8",
        "int16",
        "int32",
        "int64",
        "float16",
        "float32",
        "float64"]
    start_mem = df.memory_usage().sum() / 1024 ** 2
    for col in df.columns:
        col_type = df[col].dtypes
        if col_type in numerics:
            c_min = df[col].min()
            c_max = df[col].max()
            if str(col_type)[:3] == "int":
                if c_min > np.iinfo(
                        np.int8).min and c_max < np.iinfo(
                        np.int8).max:
                    df[col] = df[col].astype(np.int8)
                elif c_min > np.iinfo(np.int16).min and c_max < np.iinfo(np.int16).max:
                    df[col] = df[col].astype(np.int16)
                elif c_min > np.iinfo(np.int32).min and c_max < np.iinfo(np.int32).max:
                    df[col] = df[col].astype(np.int32)
                elif c_min > np.iinfo(np.int64).min and c_max < np.iinfo(np.int64).max:
                    df[col] = df[col].astype(np.int64)
            else:
                if (
                    c_min > np.finfo(np.float16).min
                    and c_max < np.finfo(np.float16).max
                ):
                    df[col] = df[col].astype(np.float16)
                elif (
                    c_min > np.finfo(np.float32).min
                    and c_max < np.finfo(np.float32).max
                ):
                    df[col] = df[col].astype(np.float32)
                else:
                    df[col] = df[col].astype(np.float64)
    end_mem = df.memory_usage().sum() / 1024 ** 2
    if verbose:
        print(
            "Mem. usage decreased to {:.2f} Mb ({:.1f}% reduction)".format(
                end_mem, 100 * (start_mem - end_mem) / start_mem
            )
        )
    return df


def get_dataframe(
        filepath: str,
        columns: List[str] = [],
        header: bool = True
    ) -> pd.DataFrame:
    """
    Get data to a dataframe.

    This function reads data from a file and transform it to a dataframe.

    Parameters:
    - filepath: str: The path to the file where is data to read.
    - columns: List[str]: A list to reaname all columns of the dataframe if the list is given.
    - header: bool: Indicates if the file has a header row. By default is True.

    Returns:
    - A pandas dataframe optimized.
    """
    # Check if the file exist
    assert path.exists(filepath), f'The file does not exist in: {filepath}'

    # Read data and transform it to a pandas dataframe
    df: pd.DataFrame = dt.fread(
        filepath,
        encoding='utf-8',
        header=header
    ).to_pandas()

    # Memory optimization of data types
    df = reduce_memory_usage(df)

    # Reanmes all columns if necessary they are given
    if len(columns):
        df.columns = columns
    return df


def parse_web_page(url: str) -> html.HtmlElement:
    """
    Parse the web page.

    This function is to get and parse the web page.

    Parameters:
    - url: str: The url of the web page to parse.

    Returns:
    - A html.HtmlElement object parsed.
    """
    response: requests = requests.get(url)

    try:
        # Evalaute the object
        response.raise_for_status()

    except requests.exceptions.RequestException as err:

        # In case of error print the error
        print(str(err))
    else:
        # Transform to text the web page
        web_page: str = response.content.decode('utf-8')

        # Parsing the web page
        return html.fromstring(web_page)
