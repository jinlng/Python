# Import the pyproj library for coordinate transformations
from pyproj import Transformer

# Create a transformer object that can convert coordinates from the Lambert-93 to WGS84 projection
transformer = Transformer.from_crs("epsg:2154", "epsg:4326")

# Define a function to convert decimal latitude values to degrees, minutes, and seconds format
def decimal_to_dms_lat(decimal_degrees):
    # Convert the decimal degree value to a positive value
    abs_degrees = abs(decimal_degrees)
    # Extract the integer degrees
    degrees = int(abs_degrees)
    # Calculate the decimal minutes
    decimal_minutes = (abs_degrees - degrees) * 60
    # Format the result as degrees and decimal minutes with N/S direction
    if decimal_degrees >= 0:
        direction = "N"
    else:
        direction = "S"
    return f"{degrees}°{decimal_minutes:.3f}'{direction}"

# Define a function to convert decimal longitude values to degrees, minutes, and seconds format
def decimal_to_dms_lon(decimal):
    d = int(decimal)
    m = int((decimal - d) * 60)
    s = (decimal - d - m/60) * 3600.0000000
    z = round(s, 4)
    if decimal >= 0:
        return f"{abs(d)}°{abs(m)}'{abs(z)}\"E"
    else:
        return f"{abs(d)}°{abs(m)}'{abs(z)}\"W"
        
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
x = 524795.419666301
y = 6809542.08749356
result = l93_to_wgs84(x, y)
print(result[0], result[1])