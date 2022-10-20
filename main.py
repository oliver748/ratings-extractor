import requests
from lxml import html
from draw import draw_spreadsheet


HEADERS = {"User-Agent": "Mozilla/5.0"}


def fetch_show(tv_show, season_num):
    url = f"https://www.imdb.com/title/{tv_show}/episodes?season={season_num}"
    fetch_webpage = requests.get(url, headers=HEADERS)
    tree = html.fromstring(fetch_webpage.content)
    return tree


def fetch_seasons_amount(tv_show):
    i = 1
    tree = fetch_show(tv_show, 1)
    while 1:
        try:
            xpath = f'//*[@id="bySeason"]/option[{i}]'
            season = tree.xpath(xpath)
            season = season[0].text
            i += 1
        except:
            return i-1


def fetch_ratings(tv_show, amount):
    all_ratings = {}
    for i in range(amount):
        print(f'Gathering season number {i+1} out of {amount}')
        tree = fetch_show(tv_show, i+1)
        season_ratings = []
        j = 0
        while 1:
            try:
                xpath = f'//*[@id="episodes_content"]/div[2]/div[2]/div[{j+1}]/div[2]/div[2]/div[1]/span[2]'
                episode_rating = tree.xpath(xpath)
                episode_rating = episode_rating[0].text
                season_ratings.append(episode_rating)
            except:
                break
            j += 1
            
        all_ratings[i+1] = season_ratings
    return all_ratings


if __name__ == "__main__":
    tv_show = input('TV Show (IMDB ID - example: tt0460681): ')
    seasons_amount = fetch_seasons_amount(tv_show)
    ratings = fetch_ratings(tv_show, seasons_amount)
    draw_spreadsheet(ratings, seasons_amount)