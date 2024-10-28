const express = require('express');
const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const { exec } = require('child_process'); // Para ejecutar comandos de shell
const app = express();
const PORT = process.env.PORT || 3000;


app.use(express.urlencoded({ extended: true }));

// Ruta para el formulario
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'views', 'index.html'));
});

// Ruta para generar el informe
app.post('/generateReport', (req, res) => {
    const { nombre, fecha, comentarios, template } = req.body;

    // Asegúrate de que la plantilla seleccionada sea válida
    if (!['template.docx', 'template2.docx'].includes(template)) {
        return res.status(400).send('Plantilla no válida.');
    }

    // Ruta de la plantilla seleccionada
    const templatePath = path.join(__dirname, template);
    console.log(`Cargando plantilla desde: ${templatePath}`); // Verifica la ruta

    try {
        // Cargar la plantilla seleccionada
        const templateContent = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(templateContent);
        const doc = new Docxtemplater(zip);

        // Insertar datos en la plantilla
        doc.setData({ nombre, fecha, comentarios });
        doc.render();

        // Guardar el archivo .docx
        const outputDocxPath = path.join(__dirname, 'generatedReports', 'report.docx');
        fs.writeFileSync(outputDocxPath, doc.getZip().generate({ type: 'nodebuffer' }));

        // Convertir el archivo .docx a PDF usando LibreOffice
        const outputPdfPath = path.join(__dirname, 'generatedReports', 'report.pdf');
        exec(`libreoffice --headless --convert-to pdf --outdir "${path.dirname(outputPdfPath)}" "${outputDocxPath}"`, (error, stdout, stderr) => {
            if (error) {
                console.error(`Error al convertir a PDF: ${stderr}`);
                return res.status(500).send('Error al generar el informe PDF.');
            }

            // Descargar el archivo PDF generado
            res.download(outputPdfPath, 'report.pdf', (err) => {
                if (err) {
                    console.error('Error al enviar el archivo:', err);
                    res.status(500).send('Error al descargar el informe');
                }
                // Eliminar archivos temporales después de la descarga
                fs.unlinkSync(outputDocxPath);
                fs.unlinkSync(outputPdfPath);
            });
        });

    } catch (error) {
        console.error('Error al generar el informe:', error);
        res.status(500).send('Error al generar el informe');
    }
});

app.listen(PORT, () => {
    console.log(`Servidor en funcionamiento en http://localhost:${PORT}`);
});
