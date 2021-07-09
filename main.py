import urllib.request
import json
import math
import matplotlib.pyplot as plt
from xltpl.writerx import BookWriter

API_URL: str = 'https://api.open-elevation.com/api/v1/lookup'


def get_elev_list(lat_list, lon_list):
    # CONSTRUCT JSON
    d_ar = [{}] * len(lat_list)
    for i in range(len(lat_list)):
        d_ar[i] = {
            "latitude": lat_list[i],
            "longitude": lon_list[i]
        }
    location = {"locations": d_ar}
    json_data = json.dumps(location, skipkeys=int).encode('utf8')

    # SEND REQUEST
    response = urllib.request.Request(API_URL, json_data, headers={'Content-Type': 'application/json'})
    fp = urllib.request.urlopen(response, timeout=10000)

    # RESPONSE PROCESSING
    res_byte = fp.read()
    res_str = res_byte.decode("utf8")
    js_str = json.loads(res_str)
    fp.close()

    response_len = len(js_str['results'])
    elev_list = []
    for j in range(response_len):
        elev = js_str['results'][j]['elevation']
        elev_list.append(elev)

    return elev_list


def get_conditional_zero_level(r_i, r_0):
    r_earth = 6379
    k_i = r_i / r_0
    return ((r_0 * r_0) / (2 * r_earth)) * k_i * (1 - k_i) * 1000


def get_conditional_zero_level_list(d_list_rev, is_zero_level):
    if is_zero_level:
        conditional_zero_level_list = []
        for j in range(len(d_list_rev)):
            conditional_zero_level_list.append(get_conditional_zero_level(d_list_rev[j], d_list_rev[-1]))
        return conditional_zero_level_list
    return [0 for _ in range(len(d_list_rev))]


# HAVERSINE FUNCTION
def haversine(lat1, lon1, lat2, lon2):
    lat1_rad = math.radians(lat1)
    lat2_rad = math.radians(lat2)
    lon1_rad = math.radians(lon1)
    lon2_rad = math.radians(lon2)
    delta_lat = lat2_rad - lat1_rad
    delta_lon = lon2_rad - lon1_rad
    a = math.sqrt(
        (math.sin(delta_lat / 2)) ** 2 + math.cos(lat1_rad) * math.cos(lat2_rad) * (math.sin(delta_lon / 2)) ** 2)
    d = 2 * 6371000 * math.asin(a)
    return d


def write_report(elev_list, d_list_rev):
    data = []
    for index, _ in enumerate(elev_list):
        data.append({
            "index": index + 1,
            "elevation": str(round(elev_list[index], 3)).replace(".", ","),
            "distance": str(round(d_list_rev[index], 3)).replace(".", ","),
        })
    writer = BookWriter("template.xlsx")
    context = {"data": data, "sheet_name": "Данные"}
    writer.render_book(payloads=[context])
    writer.save("result.xlsx")


def plot_elevation_profile(elev_list, d_list_rev, conditional_zero_level_list, elev_list_without_forest):
    # BASIC STAT INFORMATION
    mean_elev = (sum(elev_list) / len(elev_list))
    min_elev = min(elev_list)
    max_elev = max(elev_list)
    distance = d_list_rev[-1]

    base_reg = 0
    plt.figure(figsize=(10, 4))
    plt.plot(d_list_rev, elev_list_without_forest)
    plt.plot(d_list_rev, conditional_zero_level_list)
    plt.plot([0, distance], [min_elev, min_elev], '--g', label='минимум: ' + "{:.3f}".format(min_elev) + ' м')
    plt.plot([0, distance], [max_elev, max_elev], '--r', label='максимум: ' + "{:.3f}".format(max_elev) + ' м')
    plt.plot([0, distance], [mean_elev, mean_elev], '--y', label='в среднем: ' + "{:.3f}".format(mean_elev) + ' м')
    plt.fill_between(d_list_rev, elev_list_without_forest, base_reg, alpha=0.1)
    plt.fill_between(d_list_rev, elev_list, elev_list_without_forest, alpha=0.3, color="green")
    plt.fill_between(d_list_rev, conditional_zero_level_list, base_reg, alpha=0.1)
    plt.text(d_list_rev[0], elev_list[0], "")
    plt.text(d_list_rev[-1], elev_list[-1], "")
    plt.xlabel("Расстояние, км")
    plt.ylabel("Высота, м")
    ymax = max(elev_list)
    xpos = elev_list.index(ymax)
    xmax = d_list_rev[xpos]
    k = round(xmax/d_list_rev[-1], 2)
    text = "максимум x={:.3f} км, y={:.3f} м, k={:.2f}".format(xmax, ymax, k)
    bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)
    arrowprops = dict(arrowstyle="->", connectionstyle="angle,angleA=0,angleB=60")
    kw = dict(xycoords='data', textcoords="axes fraction",
              arrowprops=arrowprops, bbox=bbox_props, ha="right", va="top")
    plt.annotate(text, xy=(xmax, ymax), xytext=(0.99, 0.38), **kw)
    plt.grid()
    plt.legend(fontsize='small')
    plt.savefig('result.png', bbox_inches='tight')
    plt.show()


def get_new_elev_list(elev_list, conditional_zero_level_list):
    for j, elev in enumerate(elev_list):
        elev_list[j] = elev + conditional_zero_level_list[j]
    return elev_list


def main(point1, point2, s, is_zero_level):
    interval_lat = (point2[0] - point1[0]) / s  # interval for latitude
    interval_lon = (point2[1] - point1[1]) / s  # interval for longitude

    # SET A NEW VARIABLE FOR START POINT
    lat0 = point1[0]
    lon0 = point1[1]

    # LATITUDE AND LONGITUDE LIST
    lat_list = [lat0]
    lon_list = [lon0]

    # GENERATING POINTS
    for i in range(s):
        lat_step = lat0 + interval_lat
        lon_step = lon0 + interval_lon
        lon0 = lon_step
        lat0 = lat_step
        lat_list.append(lat_step)
        lon_list.append(lon_step)

    # DISTANCE CALCULATION
    d_list = []
    for j in range(len(lat_list)):
        lat_p = lat_list[j]
        lon_p = lon_list[j]
        dp = haversine(lat0, lon0, lat_p, lon_p) / 1000  # km
        d_list.append(dp)
    d_list_rev = d_list[::-1]  # reverse list

    elev_list = get_elev_list(lat_list, lon_list)
    conditional_zero_level_list = get_conditional_zero_level_list(d_list_rev, is_zero_level)
    elev_list = get_new_elev_list(elev_list, conditional_zero_level_list)
    elev_list_with_forest = get_elev_list_with_forest(elev_list)
    plot_elevation_profile(elev_list_with_forest, d_list_rev, conditional_zero_level_list, elev_list)
    write_report(elev_list_with_forest, d_list_rev)


def get_elev_list_with_forest(elev_list):
    elev_list_with_forest = elev_list.copy()
    for j, elev in enumerate(elev_list):
        if is_point_with_forest(j):
            elev_list_with_forest[j] += 20
    return elev_list_with_forest


def is_point_with_forest(point):
    if point in range(34, 43) or point in range(56, 86):
        return True
    return False


if __name__ == "__main__":
    point_1 = [float(x) for x in input("Введите координаты первой точки: ").split(",")]   # START POINT
    point_2 = [float(x) for x in input("Введите координаты второй точки: ").split(",")]  # END POINT
    amount = 99  # AMOUNT OF POINTS
    is_zero_level_input = input("Будет ли учитываться условный нулевой уровень (да/нет): ")
    if is_zero_level_input == "да":
        is_zero_level = True
    elif is_zero_level_input == "нет":
        is_zero_level = False
    else:
        is_zero_level = False
    main(point_1, point_2, amount, is_zero_level)

    # 52.189834, 24.374457
    # 52.4352341, 24.8846534

