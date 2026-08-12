from nicegui import ui
import pandas as pd
import io

ui.label("Movie HTML Creator").classes("text-2xl mb-4")

search_input = ui.input("Search (year, description, director, writer, cast)").classes("w-96")
ui.button("Search", on_click=lambda: run_search(search_input.value))

movies = []

async def load_csv(e):
    global movies

    # SmallFileUpload object
    uploaded = e.file

    # MUST await read()
    content = await uploaded.read()
    content = content.decode('utf-8')

    df = pd.read_csv(io.StringIO(content))
    movies = df.to_dict(orient='records')

    ui.notify(f"Loaded {len(movies)} movies")
    show_movie_grid()
    
def run_search(query):
    if not query:
        ui.notify("Enter a search term")
        return

    query = str(query).lower()
    results = []

    for movie in movies:
        # Convert all fields to strings safely
        year = str(movie.get('Year', '')).lower()
        desc = str(movie.get('Description', '')).lower()
        director = str(movie.get('Director', '')).lower()
        writer = str(movie.get('Writer', '')).lower()
        cast = str(movie.get('Cast', '')).lower()

        # Search in multiple fields
        if (query in year or
            query in desc or
            query in director or
            query in writer or
            query in cast):
            results.append(movie)

    if not results:
        ui.notify("No results found")
        return

    save_search_results(results, query)
    
def save_search_results(results, query):
    html = f"""
    <html>
    <head>
        <title>Search Results for {query}</title>
        <style>
            body {{ font-family: Arial; padding: 20px; }}
            header {{ text-align: center; }}
            .movie {{ margin-bottom: 30px; padding: 15px; border-bottom: 1px solid #ccc; }}
            a {{ color: blue; }}
        </style>
    </head>
    <body>
        <header><h1>Search Results for "{query}"</h1></header>
    """

    for movie in results:
        html += f"""
        <div class="movie">
            <h2>{movie['Title']} ({movie['Year']})</h2>
            <p><b>Director:</b> {movie['Director']}</p>
            <p><b>Writer:</b> {movie['Writer']}</p>
            <p><b>Description:</b> {movie['Description']}</p>
            <p><b>Cast:</b> {movie['Cast']}</p>
            <p><b>Link:</b> <a href="#" onclick="window.open('{movie['Link']}', '_blank')">Open TMDB</a></p>
        </div>
        """

    html += "</body></html>"

    filename = f"search_results_{query.replace(' ', '_')}.html"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(html)

    ui.notify(f"Search results saved as {filename}")
    

def show_movie_grid():
    grid.clear()

    with grid:
        ui.html("<h2>Movie Gallery</h2>")

        with ui.row().classes("flex-wrap gap-4"):
            for movie in movies:
                poster = movie.get("Image URL", "")
                title = movie.get("Title", "Untitled")

                with ui.card().classes("w-48"):
                    ui.image(poster).classes("w-full h-64 object-cover rounded")
                    ui.label(title).classes("text-center text-sm mt-2")

                    ui.button("Details", on_click=lambda m=movie: show_popup(m)) \
                        .classes("mt-2 w-full")

def show_popup(movie):
    with ui.dialog() as dialog:
        with ui.card().classes("p-4 w-96"):
            ui.label(movie.get("Title", "")).classes("text-xl mb-2")

            html = "<ul style='line-height:1.6;'>"
            html += f"<li><b>Year:</b> {movie.get('Year')}</li>"
            html += f"<li><b>Rating:</b> {movie.get('Rating')}</li>"
            html += f"<li><b>Director:</b> {movie.get('Director')}</li>"
            html += f"<li><b>Writer:</b> {movie.get('Writer')}</li>"
            html += f"<li><b>Description:</b> {movie.get('Description')}</li>"
            html += f"<li><b>Cast:</b> {movie.get('Cast')}</li>"
            html += f"<li><b>Link:</b> <a href='{movie.get('Link')}' target='_blank'>Open TMDB</a></li>"
            html += "</ul>"

            ui.html(html)
            ui.button("Close", on_click=dialog.close).classes("mt-2")

    dialog.open()
    
def save_html_file():
    html = """
    <html>
    <head>
        <title>Movie Gallery</title>
        <style>
            body { font-family: Arial; padding: 20px; background: #f5f5f5; }
            header, footer { text-align: center; padding: 10px; background: #222; color: white; }

            .search-box { text-align: center; margin-bottom: 20px; }
            .search-box input {
                width: 300px;
                padding: 8px;
                font-size: 14px;
                border-radius: 6px;
                border: 1px solid #aaa;
            }

            .grid { display: flex; flex-wrap: wrap; gap: 20px; }

            .movie {
                width: 180px;
                background: white;
                padding: 10px;
                border-radius: 10px;
                cursor: pointer;
                position: relative;
            }

            .movie img {
                width: 100%;
                height: 250px;
                object-fit: cover;
                border-radius: 8px;
            }

            .title {
                text-align: center;
                margin-top: 8px;
                font-size: 14px;
                font-weight: bold;
                position: relative;
            }

            .cast-preview {
                display: none;
                position: absolute;
                top: 20px;
                left: 50%;
                transform: translateX(-50%);
                background: white;
                padding: 10px;
                border-radius: 8px;
                box-shadow: 0 0 10px rgba(0,0,0,0.3);
                width: 160px;
                z-index: 999;
                font-size: 12px;
            }

            .title:hover .cast-preview {
                display: block;
            }

            .popup-bg {
                display: none;
                position: fixed;
                top: 0; left: 0;
                width: 100%; height: 100%;
                background: rgba(0,0,0,0.6);
                justify-content: center;
                align-items: center;
                z-index: 1000;
            }

            .popup {
                background: white;
                padding: 20px;
                width: 400px;
                border-radius: 10px;
            }

            .close-btn {
                margin-top: 10px;
                padding: 8px 12px;
                background: #222;
                color: white;
                border: none;
                cursor: pointer;
                border-radius: 5px;
            }
        </style>

        <script>
            function openPopup(id) {
                document.getElementById(id).style.display = 'flex';
            }
            function closePopup(id) {
                document.getElementById(id).style.display = 'none';
            }

            function searchMovies() {
                let q = document.getElementById('search').value.toLowerCase();
                let items = document.getElementsByClassName('movie');

                for (let item of items) {
                    let text = item.getAttribute('data-search');
                    item.style.display = text.includes(q) ? 'block' : 'none';
                }
            }
        </script>
    </head>

    <body>
        <header><h1>Movie Gallery</h1></header>

        <div class="search-box">
            <input id="search" type="text" placeholder="Search movies..."
                   onkeyup="searchMovies()">
        </div>

        <div class="grid">
    """

    for i, movie in enumerate(movies):
        popup_id = f"popup_{i}"

        searchable = (
            f"{movie['Title']} {movie['Year']} {movie['Rating']} {movie['Description']} "
            f"{movie['Director']} {movie['Writer']} {movie['Cast']}"
        ).lower()

        cast_preview = "<br>".join(movie['Cast'].split(";")[:5])

        html += f"""
        <div class="movie" data-search="{searchable}" onclick="openPopup('{popup_id}')">
            <img src="{movie['Image URL']}">

            <div class="title">
                {movie['Title']}
                <div class="cast-preview">
                    <b>Cast:</b><br>
                    {cast_preview}
                </div>
            </div>
        </div>

        <div id="{popup_id}" class="popup-bg">
            <div class="popup">
                <h2>{movie['Title']}</h2>
                <ul>
                    <li><b>Year:</b> {movie['Year']}</li>
                    <li><b>Rating:</b> {movie['Rating']}</li>
                    <li><b>Director:</b> {movie['Director']}</li>
                    <li><b>Writer:</b> {movie['Writer']}</li>
                    <li><b>Description:</b> {movie['Description']}</li>
                    <li><b>Cast:</b> {movie['Cast']}</li>
                    <li><b>Link:</b> <a href="#" onclick="window.open('{movie['Link']}', '_blank')">Open TMDB</a></li>
                </ul>
                <button class="close-btn" onclick="closePopup('{popup_id}')">Close</button>
            </div>
        </div>
        """

    html += """
        </div>
        <footer><p>Generated by NiceGUI Movie Creator</p></footer>
    </body>
    </html>
    """

    with open("movie_gallery.html", "w", encoding="utf-8") as f:
        f.write(html)

    ui.notify("Saved as movie_gallery.html")

def generate_html():
    html_output.clear()

    html = """
    <html>
    <head>
        <title>Movie Report</title>
        <style>
            body { font-family: Arial; padding: 20px; background: #f5f5f5; }
            header, footer { text-align: center; padding: 10px; background: #222; color: white; }
            .movie { margin-bottom: 40px; padding: 20px; background: white; border-radius: 10px; }
            img { width: 150px; height: 225px; object-fit: cover; border-radius: 8px; }
            ul { line-height: 1.6; }
        </style>
    </head>
    <body>
        <header><h1>Movie Report</h1></header>
    """

    for movie in movies:
        html += "<div class='movie'>"
        html += f"<h2>{movie.get('Title')}</h2>"
        html += f"<img src='{movie.get('Image URL')}' />"
        html += "<ul>"
        for key, value in movie.items():
            html += f"<li><b>{key}:</b> {value}</li>"
        html += "</ul></div>"

    html += "<footer><p>Generated by NiceGUI Movie Creator</p></footer>"
    html += "</body></html>"

    # ⭐ Correct NiceGUI usage
    with html_output:
        ui.html(html)


ui.upload(on_upload=load_csv).classes("mb-4")

grid = ui.column()
html_output = ui.column()

ui.button("Generate HTML Page", on_click=generate_html).classes("mt-4")
ui.button("Save Grid as HTML", on_click=save_html_file).classes("mt-4")


ui.run()
