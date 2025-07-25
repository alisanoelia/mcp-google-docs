import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { google, docs_v1 } from "googleapis";
import { JWT } from "google-auth-library";
import fs from "fs/promises";
import path from "path";

const server = new McpServer({
  name: "mcp-google-docs",
  version: "0.1.0",
});

const KEY_FILE_PATH = path.resolve("./service-account.json");
const SCOPES = ["https://www.googleapis.com/auth/documents"];

async function getAuthClient() {
  try {
    const keyFileContent = await fs.readFile(KEY_FILE_PATH, "utf-8");
    const { client_email, private_key } = JSON.parse(keyFileContent);
    const auth = new JWT({
      email: client_email,
      key: private_key,
      scopes: SCOPES,
    });
    return auth;
  } catch (err) {
    console.error(
      "\n[ERROR] No se pudo autenticar con la cuenta de servicio.",
      "Asegúrate de que 'service-account.json' existe y es válido.\n",
    );
    process.exit(1);
  }
}

// --- Herramienta 1: Obtener título ---
server.tool(
  "get_title",
  "Obtiene el título de un documento de Google Docs",
  {
    documentId: z.string().describe("El ID del documento de Google Docs."),
  },
  async ({ documentId }) => {
    const auth = await getAuthClient();
    const docs = google.docs({ version: "v1", auth });
    try {
      const res = await docs.documents.get({ documentId: documentId });
      return {
        content: [
          {
            type: "text",
            text: `El título del documento es: ${res.data.title}`,
          },
        ],
      };
    } catch (err: any) {
      throw new Error(`Error al obtener el documento: ${err.message}`);
    }
  },
);

// --- Herramienta 2: Reemplazar contenido ---
server.tool(
  "update_document_content",
  "Reemplaza todo el contenido de un documento de Google Docs con texto nuevo.",
  {
    documentId: z.string().describe("El ID del documento a modificar."),
    newContent: z
      .string()
      .describe("El nuevo texto que se insertará en el documento."),
  },
  async ({ documentId, newContent }) => {
    const auth = await getAuthClient();
    const docs = google.docs({ version: "v1", auth });

    try {
      const doc = await docs.documents.get({ documentId });
      const docEndIndex = doc.data.body?.content?.slice(-1)[0]?.endIndex ?? 1;

      const requests: docs_v1.Schema$Request[] = [];
      if (docEndIndex > 1) {
        requests.push({
          deleteContentRange: {
            range: { startIndex: 1, endIndex: docEndIndex - 1 },
          },
        });
      }
      requests.push({
        insertText: { location: { index: 1 }, text: newContent },
      });

      await docs.documents.batchUpdate({
        documentId: documentId,
        requestBody: { requests: requests },
      });

      return {
        content: [
          {
            type: "text",
            text: `El documento con ID ${documentId} ha sido actualizado correctamente.`,
          },
        ],
      };
    } catch (err: any) {
      throw new Error(`Error al actualizar el documento: ${err.message}`);
    }
  },
);

server.tool(
  "append_to_document",
  "Añade texto al final de un documento de Google Docs sin borrar el contenido existente.",
  {
    documentId: z.string().describe("El ID del documento a modificar."),
    textToAppend: z
      .string()
      .describe("El texto que se añadirá al final del documento."),
  },
  async ({ documentId, textToAppend }) => {
    const auth = await getAuthClient();
    const docs = google.docs({ version: "v1", auth });

    try {
      const doc = await docs.documents.get({ documentId });
      const docEndIndex = doc.data.body?.content?.slice(-1)[0]?.endIndex ?? 1;

      const requests: docs_v1.Schema$Request[] = [
        {
          insertText: {
            location: { index: docEndIndex - 1 },
            text: "\n" + textToAppend,
          },
        },
      ];

      await docs.documents.batchUpdate({
        documentId: documentId,
        requestBody: { requests: requests },
      });

      return {
        content: [
          {
            type: "text",
            text: `Se ha añadido el texto al final del documento con ID ${documentId}.`,
          },
        ],
      };
    } catch (err: any) {
      throw new Error(`Error al añadir contenido al documento: ${err.message}`);
    }
  },
);

server.tool(
  "read_document",
  "Lee y devuelve todo el contenido de texto de un documento de Google Docs.",
  {
    documentId: z.string().describe("El ID del documento a leer."),
  },
  async ({ documentId }) => {
    const auth = await getAuthClient();
    const docs = google.docs({ version: "v1", auth });

    try {
      const doc = await docs.documents.get({ documentId });
      const content = doc.data.body?.content;

      if (!content) {
        return {
          content: [{ type: "text", text: "El documento está vacío." }],
        };
      }

      const fullText = content
        .map((structuralElement) => {
          if (structuralElement.paragraph) {
            return structuralElement.paragraph.elements
              ?.map((element) => element.textRun?.content || "")
              .join("");
          }
          return "";
        })
        .join("");

      return {
        content: [
          {
            type: "text",
            text: fullText || "El documento no contiene texto visible.",
          },
        ],
      };
    } catch (err: any) {
      throw new Error(`Error al leer el documento: ${err.message}`);
    }
  },
);

server.tool(
  "format_text",
  "Aplica múltiples formatos (negrita, cursiva, color, etc.) a un texto específico.",
  {
    documentId: z.string().describe("El ID del documento a modificar."),
    textToFind: z
      .string()
      .describe("El texto exacto al que se aplicará el formato."),
    // --- Parámetros de formato opcionales ---
    bold: z.boolean().optional().describe("Aplicar o quitar negrita."),
    italic: z.boolean().optional().describe("Aplicar o quitar cursiva."),
    underline: z.boolean().optional().describe("Aplicar o quitar subrayado."),
    fontSize: z
      .number()
      .optional()
      .describe("Tamaño de la fuente en puntos (pt)."),
    foregroundColor: z
      .string()
      .optional()
      .describe("Color del texto en formato HEX (ej: '#FF0000')."),
    fontFamily: z
      .string()
      .optional()
      .describe("Fuente del texto (ej: 'Arial', 'Times New Roman')."),
  },
  async (args) => {
    const { documentId, textToFind, ...formatOptions } = args;
    const auth = await getAuthClient();
    const docs = google.docs({ version: "v1", auth });

    try {
      // Paso 1: Encontrar el texto (lógica idéntica a la anterior)
      const doc = await docs.documents.get({ documentId });
      const content = doc.data.body?.content;
      if (!content) throw new Error("Documento vacío.");

      const fullText = content
        .map(
          (el) =>
            el.paragraph?.elements
              ?.map((e) => e.textRun?.content || "")
              .join("") || "",
        )
        .join("");
      const startIndex = fullText.indexOf(textToFind);
      if (startIndex === -1)
        throw new Error(`El texto "${textToFind}" no fue encontrado.`);

      const apiStartIndex = startIndex + 1;
      const apiEndIndex = startIndex + textToFind.length + 1;

      // Paso 2: Construir dinámicamente el objeto de estilo y la máscara de campos
      const textStyle: any = {};
      const fields: string[] = [];

      if (formatOptions.bold !== undefined) {
        textStyle.bold = formatOptions.bold;
        fields.push("bold");
      }
      if (formatOptions.italic !== undefined) {
        textStyle.italic = formatOptions.italic;
        fields.push("italic");
      }
      if (formatOptions.underline !== undefined) {
        textStyle.underline = formatOptions.underline;
        fields.push("underline");
      }
      if (formatOptions.fontSize) {
        textStyle.fontSize = { magnitude: formatOptions.fontSize, unit: "PT" };
        fields.push("fontSize");
      }
      if (formatOptions.fontFamily) {
        textStyle.weightedFontFamily = { fontFamily: formatOptions.fontFamily };
        fields.push("weightedFontFamily.fontFamily");
      }
      if (formatOptions.foregroundColor) {
        const rgbColor = hexToRgb(formatOptions.foregroundColor);
        if (rgbColor) {
          textStyle.foregroundColor = { color: { rgbColor } };
          fields.push("foregroundColor");
        }
      }

      if (fields.length === 0) {
        throw new Error("No se especificó ningún formato para aplicar.");
      }

      // Paso 3: Ejecutar la petición con los estilos dinámicos
      const requests = [
        {
          updateTextStyle: {
            range: { startIndex: apiStartIndex, endIndex: apiEndIndex },
            textStyle: textStyle,
            fields: fields.join(","), // Une todos los campos a modificar
          },
        },
      ];

      await docs.documents.batchUpdate({
        documentId,
        requestBody: { requests },
      });

      return {
        content: [
          {
            type: "text",
            text: `Se aplicó el formato al texto "${textToFind}".`,
          },
        ],
      };
    } catch (err: any) {
      throw new Error(`Error al aplicar formato: ${err.message}`);
    }
  },
);

// --- Iniciar Servidor ---
const transport = new StdioServerTransport();
await server.connect(transport);
console.error(
  "Servidor MCP con cuenta de servicio iniciado. Listo para recibir solicitudes.",
);
