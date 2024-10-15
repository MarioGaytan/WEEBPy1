from flask import Flask, render_template, request, redirect, url_for
import openpyxl
import os


app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/contact')
def contact():
    return render_template('contact.html')

@app.route('/computadoras')
def computadoras():
    return render_template('computadoras.html')

# Ruta para manejar el formulario de nombre y contraseña
@app.route('/submit_form', methods=['POST'])
def submit_form():
    nombre = request.form.get('nombre')
    password = request.form.get('password')
    
    # Llamar a la función que guarda los datos en un archivo Excel
    guardar_en_excel(nombre, password)
    
    # Redirigir a la página de elementos
    return redirect(url_for('elements'))

# Nueva ruta para la página de elementos
@app.route('/elements')
def elements():
    return render_template('elements.html')

def guardar_en_excel(nombre, password):
    archivo_excel = 'datos_usuarios.xlsx'
    
    # Verificar si el archivo ya existe
    if not os.path.exists(archivo_excel):
        # Si el archivo no existe, crearlo y agregar encabezados
        libro = openpyxl.Workbook()
        hoja = libro.active
        hoja.title = "Usuarios"
        hoja.append(["Nombre", "Contraseña"])  # Encabezados
    else:
        # Si el archivo existe, cargarlo
        libro = openpyxl.load_workbook(archivo_excel)
        hoja = libro.active

    # Agregar los datos del formulario a una nueva fila
    hoja.append([nombre, password])
    
    # Guardar el archivo
    libro.save(archivo_excel)

if __name__ == '__main__':
    app.run(debug=True)
