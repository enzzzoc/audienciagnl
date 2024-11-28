document.getElementById('archivoExcel').addEventListener('change', function (event) {
    const archivo = event.target.files[0];

    if (archivo) {
        const lector = new FileReader();
        lector.onload = function (e) {
            const datos = new Uint8Array(e.target.result);
            const workbook = XLSX.read(datos, { type: 'array' });
            const hojaNombre = workbook.SheetNames[0];
            const hoja = workbook.Sheets[hojaNombre];
            const datosJSON = XLSX.utils.sheet_to_json(hoja, { header: 1 });

            mostrarDatos(datosJSON);
        };
        lector.readAsArrayBuffer(archivo);
    }
});

function mostrarDatos(datos) {
    const tabla = document.getElementById('tablaAsistencia');
    tabla.innerHTML = '';

    datos.slice(1).forEach((fila, index) => {
        const [apellido, nombre, dni] = fila;

        if (apellido && nombre && dni) {
            const filaHTML = `
                <tr>
                    <td>${index + 1}</td>
                    <td>${apellido}</td>
                    <td>${nombre}</td>
                    <td>${dni}</td>
                    <td>
                        <label>
                            <input type="radio" name="estado-${index}" value="Presente" /> Presente
                        </label>
                        <label>
                            <input type="radio" name="estado-${index}" value="Ausente" /> Ausente
                        </label>
                    </td>
                </tr>
            `;
            tabla.insertAdjacentHTML('beforeend', filaHTML);
        }
    });
}

function filtrarPorDNI() {
    const dniBuscado = document.getElementById('buscarDni').value.trim();
    const filas = document.querySelectorAll('#tablaAsistencia tr');

    filas.forEach(fila => {
        const dniCelda = fila.cells[3]?.textContent || '';
        fila.style.display = dniCelda === dniBuscado ? '' : 'none';
    });
}

function restablecerTabla() {
    const filas = document.querySelectorAll('#tablaAsistencia tr');
    filas.forEach(fila => (fila.style.display = ''));
}

document.getElementById('descargarExcel').addEventListener('click', function () {
    const tabla = document.getElementById('tablaAsistencia');
    const filas = tabla.querySelectorAll('tr');

    const datosExcel = [['Apellido', 'Nombre', 'DNI', 'Estado']];
    filas.forEach((fila, index) => {
        if (fila.style.display !== 'none') {
            const celdas = fila.querySelectorAll('td');
            const apellido = celdas[1].textContent;
            const nombre = celdas[2].textContent;
            const dni = celdas[3].textContent;
            const estado = Array.from(celdas[4].querySelectorAll('input')).find(input => input.checked)?.value || '';

            if (estado) {
                datosExcel.push([apellido, nombre, dni, estado]);
            }
        }
    });

    const hoja = XLSX.utils.aoa_to_sheet(datosExcel);
    const libro = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(libro, hoja, 'Asistencia');

    XLSX.writeFile(libro, 'asistencia.xlsx');
});

