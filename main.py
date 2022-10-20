import requests
from lxml import html
from draw import draw_spreadsheet
import time


HEADERS = {"User-Agent": "Mozilla/5.0"}
OUTPUT_INFO = True # controls if the process should be printed
SAFETY = True # should be kept on; makes process take longer but ensures all ratings are retrieved
SLEEP_TIME = 0.75 # time it sleeps with safety enabled - 1.5 secs recommended

def fetch_show(tv_show, season_num):
    url = f"https://www.imdb.com/title/{tv_show}/episodes?season={season_num}"
    fetch_webpage = requests.get(url, headers=HEADERS)
    tree = html.fromstring(fetch_webpage.content)
    return tree


def fetch_seasons_amount(tv_show):
    tree = fetch_show(tv_show, 1)
    count = 1
    while True:
        try:
            xpath = f'//*[@id="bySeason"]/option[{count}]'
            season = tree.xpath(xpath)
            season = season[0].text
            count += 1
        except:
            return count-1


def fetch_ratings(tv_show, amount):
    all_ratings = {} # all ratings from every episode
    for i in range(amount):
        print(f'Gathering season number {i+1} out of {amount}')
        tree = fetch_show(tv_show, i+1)
        season_ratings = [] # all ratings from every episode in specific season
        count = 1
        while True: # finds rating for each episode - continues until it cant find more episodes
            try:
                xpath = f'//*[@id="episodes_content"]/div[2]/div[2]/div[{count}]/div[2]/div[2]/div[1]/span[2]'
                episode_rating = tree.xpath(xpath)
                episode_rating = episode_rating[0].text
                season_ratings.append(episode_rating)
            except:
                break
            count += 1
        all_ratings[i+1] = season_ratings
        
        if SAFETY: 
            time.sleep(SLEEP_TIME)
            
    return all_ratings


if __name__ == "__main__":
    print('NOTE: Close spreadsheet if opened\n')
    tv_show = input('TV Show (IMDB ID - example: tt0460681): ')
    seasons_amount = fetch_seasons_amount(tv_show)
    ratings = fetch_ratings(tv_show, seasons_amount)
    draw_spreadsheet(ratings, seasons_amount)
