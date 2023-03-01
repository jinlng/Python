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


# Example usage of the functions with a single set of coordinates
x = 491966.081673669
y = 6823056.68796353
result = l93_to_wgs84(x, y)
print(result[0], result[1])