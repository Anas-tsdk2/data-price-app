from flask import Flask, request, jsonify
from flask_cors import CORS  # Import de CORS

app = Flask(__name__)
CORS(app)  # Autorise toutes les origines par défaut

@app.route('/process_csv', methods=['POST'])
def process_csv():
    # Votre logique de traitement CSV ici
    return jsonify({"message": "Traitement réussi"})

if __name__ == '__main__':
    app.run(debug=True)
