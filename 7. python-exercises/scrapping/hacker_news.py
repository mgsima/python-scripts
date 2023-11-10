import requests
from bs4 import BeautifulSoup
import pprint



def get_pages(link):
    res = requests.get(link)
    soup = BeautifulSoup(res.text, 'html.parser')    # recibimos una string con requests y bs4 nos permite convertir eso en datos usables
    
    # Buscando los datos que nos interesan
    links = soup.select('.titleline')
    subtext = soup.select('.subtext')
    return create_custome_hn(links, subtext)


def sort_stories_by_votes(hnlist):
    return sorted(hnlist, key=lambda k:k['Votes'], reverse=True)

def create_custome_hn(links, votes):
    hn = []
    for i, item in enumerate(links):
        title = links[i].getText()
        href = links[i].select_one('a')['href']
        vote = votes[i].select('.score')
        if len(vote):
            points = int(vote[0].getText().replace(' points', ''))
            if points > 99:
                hn.append({'title': title, 'link': href, 'Votes': points})

    return hn

pages = 3
results = []
for i in range(pages):
    print(i)
    link = f'https://news.ycombinator.com/?p={i+1}'
    page_data = get_pages(link) 
    results.extend(page_data)
r = sort_stories_by_votes(results)


pprint.pprint(r)