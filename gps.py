# Import the pyproj library for coordinate transformations
from pyproj import Transformer

# Create a transformer object that can convert coordinates from the Lambert-93 to WGS84 projection
transformer = Transformer.from_crs("epsg:2154", "epsg:4326")

# Define a function to convert decimal latitude values to degrees, minutes, and seconds format
def decimal_to_dms_lat(decimal):
    d = int(decimal)
    m = int((decimal - d) * 60)
    s = (decimal - d - m/60) * 3600.0000000
    z = round(s, 4)
    if decimal >= 0:
        return f"{abs(d)}°{abs(m)}'{abs(z)}\"N"
    else:
        return f"{abs(d)}°{abs(m)}'{abs(z)}\"S"

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
    xx, yy = transformer.transform(x, y)
    # Convert the decimal latitude and longitude values to degrees, minutes, and seconds format
    latitude = decimal_to_dms_lat(xx)
    longitude = decimal_to_dms_lon(yy)
    # Print the resulting latitude and longitude values
    print(f"Latitude: {latitude}, Longitude: {longitude}")

# Example usage of the functions with a single set of coordinates
x = 524795.419666301
y = 6809542.08749356
l93_to_wgs84(x, y)

