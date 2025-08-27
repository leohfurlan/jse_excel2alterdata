import os
import uuid
import shutil
import time
import threading
from flask import Flask, request, render_template, send_from_directory, url_for
from werkzeug.utils import secure_filename

# Importa a função principal do nosso script original
import excel2alterdata

# Configuração inicial do Flask
app = Flask(__name__)
os.makedirs("temp_uploads", exist_ok=True)
os.makedirs("temp_outputs", exist_ok=True)
app.config['UPLOAD_FOLDER'] = 'temp_uploads'
app.config['OUTPUT_FOLDER'] = 'temp_outputs'
app.config['PERMANENT_SESSION_LIFETIME'] = 600 # 10 minutos

def cleanup_folder(folder_path):
    """Espera 10 minutos e depois deleta a pasta especificada."""
    try:
        time.sleep(600)  # 600 segundos = 10 minutos
        shutil.rmtree(folder_path)
        print(f"Pasta de sessão limpa com sucesso: {folder_path}")
    except Exception as e:
        print(f"Erro ao limpar a pasta {folder_path}: {e}")

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        session_id = str(uuid.uuid4())
        upload_path = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], session_id)
        os.makedirs(upload_path, exist_ok=True)
        os.makedirs(output_path, exist_ok=True)

        files = request.files.getlist('files')
        
        if not files or files[0].filename == '':
            return render_template('index.html', error="Nenhum arquivo selecionado. Por favor, escolha uma ou mais planilhas.")

        for file in files:
            if file:
                filename = secure_filename(file.filename)
                file.save(os.path.join(upload_path, filename))
        
        # Chama a lógica de processamento, que agora retorna o resumo e os erros
        summary, errors_df = excel2alterdata.main(
            in_dir=upload_path, 
            out_dir=output_path, 
            mapping="config/mapping.yaml"
        )

        # Prepara os links para download
        download_files = {}
        if summary:
            for key, path in summary.get("saidas", {}).items():
                filename = os.path.basename(path)
                if os.path.exists(path):
                    download_files[key] = url_for('download_file', session_id=session_id, filename=filename)
        
        # Prepara as inconsistências para exibição na tela
        inconsistencies = []
        if not errors_df.empty:
            inconsistencies = errors_df.to_dict('records')

        # Limpa a pasta de upload imediatamente
        shutil.rmtree(upload_path)
        
        # Agenda a limpeza da pasta de saída para daqui a 10 minutos
        cleanup_thread = threading.Thread(target=cleanup_folder, args=(output_path,))
        cleanup_thread.start()

        return render_template('index.html', summary=summary, download_files=download_files, inconsistencies=inconsistencies)

    return render_template('index.html')

@app.route('/download/<session_id>/<filename>')
def download_file(session_id, filename):
    directory = os.path.join(app.config['OUTPUT_FOLDER'], session_id)
    return send_from_directory(directory, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
