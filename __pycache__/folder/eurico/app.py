from flask import Flask, jsonify, request
from mongo_client import get_mongo_connection
from omdb_client import fetch_movie

app = Flask(__name__)
movies_collection = get_mongo_connection()

@app.route("/movies", methods=["GET"])
def get_movies():
    """
    Retorna todos os filmes armazenados no MongoDB.
    """
    movies = list(movies_collection.find({}, {"_id": 0}))
    return jsonify(movies)

@app.route("/movies", methods=["POST"])
def add_movie():
    """
    Adiciona um novo filme ao MongoDB usando a OMDb API.
    """
    data = request.get_json()
    title = data.get("title")
    
    if not title:
        return jsonify({"error": "O campo 'title' é obrigatório."}), 400
    
    movie = fetch_movie(title)
    
    if not movie:
        return jsonify({"error": "Filme não encontrado na OMDb API."}), 404
    
    movies_collection.insert_one(movie)
    
    return jsonify({"message": "Filme adicionado com sucesso!", "movie": movie})

if __name__ == "__main__":
    app.run(debug=True, port=5000, host="0.0.0.0")