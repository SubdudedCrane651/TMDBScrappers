import requests
from bs4 import BeautifulSoup

def get_movie_details_tmdb(title):
    search_url = f"https://www.themoviedb.org/search?query={title.replace(' ', '%20')}"
    response = requests.get(search_url)
    soup = BeautifulSoup(response.text, 'html.parser')

    # Find the first result's link
    first_result = soup.find('a', class_='result')
    if first_result:
        movie_link = "https://www.themoviedb.org" + first_result['href']
        
        # Get movie details from the movie page
        movie_response = requests.get(movie_link)
        movie_soup = BeautifulSoup(movie_response.text, 'html.parser')

        movie_title = movie_soup.find('h2').text.strip()
        overview = movie_soup.find('div', class_='overview').text.strip() if movie_soup.find('div', class_='overview') else 'N/A'
        rating_value = movie_soup.find('span', class_='user_score_chart')['data-percent'] if movie_soup.find('span', class_='user_score_chart') else 'N/A'
        
                # Get cast information
        #cast_list = movie_soup.select('div.cast_scroller li.card')
        cast_list = movie_soup.select('ol.people li.card')
        cast = []
        for actor in cast_list:
             actor_name_tag = actor.find('p').find('a')
             actor_name = actor_name_tag.text.strip() if actor_name_tag else 'N/A'
             character = actor.find('p', class_='character').text.strip()
             cast.append({'actor_name': actor_name, 'character': character}) 

        # Find the director and writer names
        people_section = movie_soup.find('ol', class_='people no_image')
        director = 'N/A'
        writer = 'N/A'
        
        if people_section:
            profiles = people_section.find_all('li', class_='profile')
            for profile in profiles:
                role = profile.find('p', class_='character').text.strip()
                name = profile.find('a').text.strip()
                if 'Director' in role:
                    director = name
                if 'Writer' in role:
                    writer = name

        return {
            'Title': movie_title,
            'Rating': rating_value,
            'Overview': overview,
            'Cast': cast,
            'Director': director,
            'Writer': writer,
            'URL': movie_link
        }
    else:
        return None

# Example usage
movie_title = "Interstellar"
movie_details = get_movie_details_tmdb(movie_title)
if movie_details:
    print(f"Title: {movie_details['Title']}")
    print(f"Overview: {movie_details['Overview']}")
    print(f"Rating: {movie_details['Rating']}")
    print("\nCharacters:")
    for member in movie_details['Cast']:
        print(f"{member['actor_name']} / {member['character']}")
    print(f"Director: {movie_details['Director']}")
    print(f"Writer: {movie_details['Writer']}")
    print(f"URL: {movie_details['URL']}")
else:
    print("Movie not found")
