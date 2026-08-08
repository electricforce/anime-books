import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle } from "docx";
import fs from "fs";
import path from "path";

// Definición de colores elegantes
const COLOR_PRIMARY = "4F46E5";   // Indigo (Acentos y títulos)
const COLOR_SECONDARY = "111827"; // Negro/Gris muy oscuro (Título principal)
const COLOR_DARK = "374151";      // Gris carbón (Texto principal)
const COLOR_LIGHT = "6B7280";     // Gris medio (Contacto, metadatos)

// Estilo del borde inferior para las secciones
const sectionBorderBottom = {
    color: COLOR_PRIMARY,
    space: 4,
    value: BorderStyle.SINGLE,
    size: 8, // 1 pt
};

// Helper para crear títulos de sección con borde inferior
function createSectionHeader(title) {
    return new Paragraph({
        text: title,
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 240, after: 120 },
        border: {
            bottom: sectionBorderBottom
        },
    });
}

// Helper para viñetas resumidas
function createBulletPoint(boldPrefix, text) {
    return new Paragraph({
        bullet: {
            level: 0
        },
        spacing: { before: 40, after: 40 },
        children: [
            new TextRun({
                text: `${boldPrefix}: `,
                bold: true,
                color: COLOR_PRIMARY,
            }),
            new TextRun({
                text: text,
                color: COLOR_DARK,
            }),
        ],
    });
}

// Estructura del Documento
const doc = new Document({
    title: "CV - Adrián Romero",
    description: "Currículum Vitae resumido enfocado en habilidades blandas y proyectos web.",
    styles: {
        default: {
            document: {
                run: {
                    font: "Segoe UI",
                    color: COLOR_DARK,
                    size: 21, // 10.5 pt para mayor compactación
                },
                paragraph: {
                    spacing: {
                        line: 240, // Interlineado 1.0 (compacto)
                    },
                },
            },
            heading1: {
                run: {
                    font: "Segoe UI",
                    color: COLOR_SECONDARY,
                    bold: true,
                    size: 40, // 20 pt
                },
                paragraph: {
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 60 },
                },
            },
            heading2: {
                run: {
                    font: "Segoe UI",
                    color: COLOR_PRIMARY,
                    bold: true,
                    size: 24, // 12 pt
                },
            },
        },
    },
    sections: [{
        properties: {
            // Márgenes moderados para que quepa todo perfectamente en una página
            page: {
                margin: {
                    top: 1000,
                    bottom: 1000,
                    left: 1000,
                    right: 1000,
                }
            }
        },
        children: [
            // ==========================================
            // ENCABEZADO Y CONTACTO
            // ==========================================
            new Paragraph({
                text: "ADRIÁN ROMERO",
                heading: HeadingLevel.HEADING_1,
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 },
                children: [
                    new TextRun({
                        text: "Desarrollador Web Full Stack",
                        bold: true,
                        color: COLOR_PRIMARY,
                        size: 22, // 11 pt
                    }),
                ],
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 },
                children: [
                    new TextRun({
                        text: "adriro1303@gmail.com  |  github.com/electricforce  |  Repositorio local: c/repositorio",
                        color: COLOR_LIGHT,
                        size: 18, // 9 pt
                    }),
                ],
            }),

            // ==========================================
            // RESUMEN PROFESIONAL
            // ==========================================
            createSectionHeader("PERFIL"),
            new Paragraph({
                alignment: AlignmentType.JUSTIFY,
                spacing: { after: 100 },
                children: [
                    new TextRun({
                        text: "Desarrollador Web Full Stack con experiencia práctica en la creación de aplicaciones dinámicas (Astro, React) y en el diseño e integración de APIs comerciales y bases de datos relacionales. Destaco por mi ",
                    }),
                    new TextRun({
                        text: "adaptabilidad técnica",
                        bold: true,
                    }),
                    new TextRun({
                        text: ", mi enfoque en la ",
                    }),
                    new TextRun({
                        text: "resolución de problemas",
                        bold: true,
                    }),
                    new TextRun({
                        text: " y mi capacidad para colaborar de forma efectiva en la integración de flujos frontend y backend.",
                    }),
                ],
            }),

            // ==========================================
            // HABILIDADES BLANDAS EN ACCIÓN
            // ==========================================
            createSectionHeader("HABILIDADES BLANDAS EN ACCIÓN"),
            createBulletPoint(
                "Colaboración y Comunicación",
                "Integración exitosa de capas de frontend y backend trabajando con APIs RESTful complejas y bases de datos para proyectos multidisciplinares como SIMA-web."
            ),
            createBulletPoint(
                "Resolución de Problemas",
                "Diseño óptimo de diagramas entidad-relación y normalización de bases de datos para sistemas transaccionales comerciales, mejorando la coherencia y acceso a los datos."
            ),
            createBulletPoint(
                "Adaptabilidad e Iniciativa",
                "Rápida asimilación y puesta en marcha de herramientas y frameworks modernos (como Astro y Tailwind CSS v4) en plazos cortos para proyectos interactivos y portales web."
            ),

            // ==========================================
            // PROYECTOS DESTACADOS
            // ==========================================
            createSectionHeader("PROYECTOS Y EXPERIENCIA"),
            
            new Paragraph({
                spacing: { before: 100, after: 60 },
                children: [
                    new TextRun({ text: "Desarrollador de Software (Proyectos Clave)  ", bold: true, color: COLOR_SECONDARY }),
                    new TextRun({ text: "—  2023 - Presente", color: COLOR_LIGHT, size: 20 })
                ]
            }),
            createBulletPoint(
                "SIMA-web & Proyecto SIMA",
                "Portal web de gestión organizativa y académica. Diseñé interfaces dinámicas en el frontend y gestioné su conexión con los servicios de datos del backend."
            ),
            createBulletPoint(
                "FarmaExpress Kiosco",
                "Aplicación comercial interactiva de consulta y ventas para farmacias, integrando consultas rápidas de base de datos."
            ),
            createBulletPoint(
                "ElectroShop & Tienda API",
                "Diseño y desarrollo de servicios web backend (APIs RESTful) con Node.js y Express para la automatización de flujos de compra y consulta."
            ),
            createBulletPoint(
                "Anime-Books Portal",
                "Catálogo de mangas y libros optimizado mediante renderizado estático y componentes reactivos con Astro y Tailwind CSS v4."
            ),

            // ==========================================
            // HABILIDADES TÉCNICAS
            // ==========================================
            createSectionHeader("HABILIDADES TÉCNICAS"),
            new Paragraph({
                spacing: { before: 60, after: 60 },
                children: [
                    new TextRun({ text: "Frontend: ", bold: true, color: COLOR_PRIMARY }),
                    new TextRun({ text: "JavaScript (ES6+), HTML5, CSS3, Astro, React, Tailwind CSS" })
                ]
            }),
            new Paragraph({
                spacing: { before: 60, after: 60 },
                children: [
                    new TextRun({ text: "Backend y BD: ", bold: true, color: COLOR_PRIMARY }),
                    new TextRun({ text: "Node.js, Express, SQL / Modelado de Bases de Datos, Git" })
                ]
            }),

            // ==========================================
            // EDUCACIÓN
            // ==========================================
            createSectionHeader("FORMACIÓN"),
            new Paragraph({
                spacing: { before: 80, after: 40 },
                children: [
                    new TextRun({ text: "Desarrollo de Software / Ingeniería de Sistemas  ", bold: true, color: COLOR_SECONDARY }),
                    new TextRun({ text: "— En curso / Egresado", color: COLOR_DARK })
                ]
            })
        ],
    }],
});

// Guardar el documento
const outputFilename = "CV_Lider_Desarrollo_Software.docx";
const outputPath = path.resolve(process.cwd(), outputFilename);

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(outputPath, buffer);
    console.log(`¡Documento CV personalizado y resumido generado con éxito en: ${outputPath}!`);
}).catch((error) => {
    console.error("Error al generar el documento Word:", error);
    process.exit(1);
});
