import { Document, Packer, Paragraph, ShadingType, Table, TableCell, TableRow, TextRun, WidthType } from "docx";
import { readFileSync, writeFileSync } from 'fs';
import * as readline from 'node:readline/promises';
import { stdin as input, stdout as output } from 'node:process';
/**
 * @typedef {import('asana').resources.Tasks.Type} AsanaTask
 */

const colorCompletion = process.argv.includes("--color-completion");
const locale = 'ES'; // Default date locale is Spanish

const headerShading = {
    fill: "#42c5f4",
    type: ShadingType.CLEAR,
};

const notCompletedShading = {
    fill: "#f57c7c",
    type: ShadingType.CLEAR,
};

const completedShading = {
    fill: "#42f468",
    type: ShadingType.CLEAR,
};

const grayShading = {
    fill: "#f6f6f6",
    type: ShadingType.CLEAR,
};


let obj;
try {
    obj = JSON.parse(readFileSync('tasks.json', 'utf8'));
} catch (error) {
    console.error("Error al leer fichero de tareas. Comprueba que existe y se llama 'tasks.json'");
    process.exit(1);
}

/**
 * @type {AsanaTask[]}
 */
const tasks = obj.data;

/**
 * @type {TableRow[]}
 */
const rows = [new TableRow({
    children: [
        new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: "Tarea", bold: true })] })],
            shading: headerShading,
        }),
        new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: "Responsable", bold: true })] })],
            shading: headerShading,
        }),
        new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: "Colaboradores", bold: true })] })],
            shading: headerShading,
        }),
        new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: "Descripción", bold: true })] })],
            shading: headerShading,
        }),
        new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: "Tiempo estimado", bold: true })] })],
            shading: headerShading,
        }),
        new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: "Fecha de completado", bold: true })] })],
            shading: headerShading,
        }),
        new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: "Estado", bold: true })] })],
            shading: headerShading,
        }),
    ],
})];

// Asks for date range
const rl = readline.createInterface({ input, output });
const initialDateStr = await rl.question('Minimum date of creation of tasks (dd-mm-yyyy) (leave blank to not use this filter): ');
const endDateStr = await rl.question('Max date of creation of tasks (dd-mm-yyyy) (leave blank to not use this filter): ');
rl.close();

let initialDate = new Date(initialDateStr.split('-').reverse().join('-')) ?? new Date();
let endDate = new Date(endDateStr.split('-').reverse().join('-'));

if (isNaN(initialDate.getTime())) initialDate = new Date(0); // If not valid, set to minimum date
if (isNaN(endDate.getTime())) endDate = new Date(); // If not valid, set to current date

let paintGray = false;
tasks.filter((task) => {
    const createdAt = new Date(task.created_at ?? "");
    return createdAt >= initialDate && createdAt <= endDate;
}).toSorted((a, b) => {
    const statusA = a.memberships[0].section?.name ?? "";
    const statusB = b.memberships[0].section?.name ?? "";
    return statusA.localeCompare(statusB);
})
    .forEach((task) => {
        const row = new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph(task.name ?? "")],
                    shading: paintGray ? grayShading : undefined,
                }),
                new TableCell({
                    children: [new Paragraph(task.assignee?.name ?? "Nadie")],
                    shading: paintGray ? grayShading : undefined,
                }),
                new TableCell({
                    children: [new Paragraph(task.followers?.map(follower => follower.name).join(", ") ?? "")],
                    shading: paintGray ? grayShading : undefined,
                }),
                new TableCell({
                    children: [new Paragraph(task.notes ?? "")],
                    shading: paintGray ? grayShading : undefined,
                }),
                new TableCell({
                    children: [new Paragraph("")],
                    shading: paintGray ? grayShading : undefined,
                }),
                new TableCell({
                    children: [new Paragraph(task.completed_at ? (new Date(task.completed_at)).toLocaleDateString(new Intl.Locale(locale)) : "No completada")],
                    shading: colorCompletion ? (task.completed ? completedShading : notCompletedShading) : paintGray ? grayShading : undefined,
                }),
                new TableCell({
                    children: [new Paragraph(task.memberships[0].section?.name ?? "")],
                    shading: paintGray ? grayShading : undefined,
                }),
            ],
        });
        rows.push(row);
        paintGray = !paintGray;
    });

const doc = new Document({
    sections: [
        {
            children: [
                new Table({
                    width: {
                        size: 9000,
                        type: WidthType.DXA,
                    },
                    rows: rows,
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    writeFileSync(`asanaTasksTable${Date.now()}.docx`, buffer);
});