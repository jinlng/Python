"""
Title: Coordinate Conversion Utility
Description:
    This code utilizes the pyproj library to perform coordinate conversion from the Lambert-93 projection to WGS84.
    However, if the sole requirement is to convert WGS84 decimal latitude and longitude values to the degrees, minutes, and seconds format,
    you can easily achieve this by making a straightforward modification in the running script.
    For instance, in the script "Création_C6.py" or "prep_C6.py",
    you can replace the existing function "l93_to_wgs84" with "wgs84_decimal_to_degree" to accomplish the desired conversion.
Author: Jingyi LIANG
Date: May 16, 2023
License: This code is the property of Jingyi LIANG. Unauthorized use or distribution is strictly prohibited.
"""

# Import the pyproj library for coordinate transformations
from pyproj import Transformer

# Create a transformer object that can convert coordinates from the Lambert-93 to WGS84 projection
transformer = Transformer.from_crs("epsg:2154", "epsg:4326")


# Define a function to convert decimal latitude/longitude values to degrees, minutes, and seconds format
def decimal_to_dms_lat(latitude):
    abs_latitude = abs(latitude)
    degrees = int(abs_latitude)
    minutes = int((abs_latitude - degrees) * 60)
    seconds = round((abs_latitude - degrees - minutes / 60) * 3600, 4)
    direction = "N" if latitude >= 0 else "S"
    return f"{degrees}°{minutes:02d}'{seconds:.4f}\"{direction}"


def decimal_to_dms_lon(longitude):
    abs_longitude = abs(longitude)
    degrees = int(abs_longitude)
    minutes = int((abs_longitude - degrees) * 60)
    seconds = round((abs_longitude - degrees - minutes / 60) * 3600, 4)
    direction = "E" if longitude >= 0 else "W"
    return f"{degrees}°{minutes:02d}'{seconds:.4f}\"{direction}"


# Define a function to convert Lambert-93 coordinates to WGS84 latitude and longitude values
def l93_to_wgs84(x, y):
    # Use the transformer object to convert the coordinates from Lambert-93 to WGS84
    result = transformer.transform(x, y)
    if result is not None:
        xx, yy = result
        # Convert the decimal latitude and longitude values to degrees, minutes, and seconds format
        latitude = decimal_to_dms_lat(xx)
        longitude = decimal_to_dms_lon(yy)
        # Return the resulting latitude and longitude values
        return latitude, longitude
    else:
        # If the transformer method returns None, return None for both latitude and longitude
        return None, None

# Define a function only to convert the decimal latitude and longitude values to degrees, minutes, and seconds format
def wgs84_decimal_to_degree(xx,yy):
    result = (xx, yy)
    if result is not None:
        xx, yy = result
        latitude = decimal_to_dms_lat(xx)
        longitude = decimal_to_dms_lon(yy)
        # Return the resulting latitude and longitude values
        return latitude, longitude
    else:
        # If the transformer method returns None, return None for both latitude and longitude
        return None, None

# Example usage of the functions with a single set of coordinates
# x = 488193.697684623
# y = 6835737.74602665
# result = l93_to_wgs84(x, y)
# print(result[0], result[1])