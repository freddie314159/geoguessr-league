type CellValue = string | number | boolean | null;

type BlockInfo = {
    name: string;
    rowIndex: number;
    colIndex: number;
};

type TableRow = {
    rank: number;
    previousRank: number | null;
    change: string;
    player: string;
    streak: string;
    roundsPlayed: CellValue;
    value: CellValue;
    form: CellValue[];
};

type ExportData = {
    freezeDate: string;
    dates: string[];
    players: string[];
    metricOrder: string[];
    metricLabels: Record<string, string>;
    tables: Record<string, TableRow[]>;
    history: Record<string, Record<string, CellValue[]>>;
};

const DATA_ROW_OFFSET = 1;
const DATA_COL_OFFSET = 0;

const HIDDEN_METRICS = new Set(["DayAvgHist"]);

const METRIC_LABELS: Record<string, string> = {
    TotScoreHist: "Total Score",
    AvgScoreHist: "Average Score",
    TodayScoreHist: "Today's Score",
    HighScoreHist: "Best Round",
    LowScoreHist: "Worst Round",
    RangeHist: "Range",
    StdDevHist: "Std Deviation",
    SlopeHist: "Slope",
    HighFallOffHist: "Biggest Fall-off",
    HighLockInHist: "Biggest Lock-in",
    AvgRBRDifHist: "Avg RbR Dif",
    AvgDistToAvgHist: "Avg Dist to Avg",
    HighStreakHist: "Longest Streak",
};

const TABLE_METRICS = Object.keys(METRIC_LABELS);

function main(workbook: ExcelScript.Workbook): string {
    const calc = workbook.getWorksheet("Calc");
    const used = calc.getUsedRange();

    if (!used) {
        throw new Error("Calc sheet is empty.");
    }

    const allValues = used.getValues() as CellValue[][];
    const blocks = findHistoryBlocks(allValues);

    const players = readNamedStringList(workbook, "Players");
    const activeFlags = readNamedBooleanList(workbook, "ActiveFlags");
    const allDates = readNamedStringList(workbook, "ScoreDates", true);

    const freezeDate = getFreezeDateIso(workbook);
    const freezeIndex = allDates.indexOf(freezeDate);

    if (freezeIndex < 0) {
        throw new Error(`FreezeDate ${freezeDate} was not found in ScoreDates.`);
    }

    const dates = allDates.slice(0, freezeIndex + 1);
    const playerCount = Math.min(players.length, activeFlags.length);

    const matrices: Record<string, CellValue[][]> = {};
    for (const block of blocks) {
        matrices[block.name] = readMatrix(
            allValues,
            block.rowIndex,
            block.colIndex,
            players.length,
            dates.length
        );
    }

    const history = buildHistoryExport(matrices, players, dates.length);
    const tables = buildAllTables(matrices, players, activeFlags, dates.length);

    const data: ExportData = {
        freezeDate,
        dates,
        players: players.slice(0, playerCount),
        metricOrder: TABLE_METRICS,
        metricLabels: METRIC_LABELS,
        tables,
        history,
    };

    return JSON.stringify(data, null, 2);
}

function findHistoryBlocks(values: CellValue[][]): BlockInfo[] {
    const blocks: BlockInfo[] = [];

    for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
            const cell = values[r][c];
            if (typeof cell === "string" && cell.endsWith("Hist") && !HIDDEN_METRICS.has(cell)) {
                blocks.push({
                    name: cell,
                    rowIndex: r,
                    colIndex: c,
                });
            }
        }
    }

    return blocks;
}

function readNamedStringList(
    workbook: ExcelScript.Workbook,
    name: string,
    treatNumbersAsDates: boolean = false
): string[] {
    const namedItem = workbook.getNamedItem(name);
    const arrayValues = namedItem.getArrayValues().getValues() as CellValue[][];

    const result: string[] = [];
    for (const row of arrayValues) {
        for (const cell of row) {
            const value = treatNumbersAsDates ? normalizeDateCell(cell) : normalizeTextCell(cell);
            if (value !== "") {
                result.push(value);
            }
        }
    }

    return result;
}

function readNamedBooleanList(workbook: ExcelScript.Workbook, name: string): boolean[] {
    const namedItem = workbook.getNamedItem(name);
    const arrayValues = namedItem.getArrayValues().getValues() as CellValue[][];

    const result: boolean[] = [];
    for (const row of arrayValues) {
        for (const cell of row) {
            if (typeof cell === "boolean") {
                result.push(cell);
            } else if (typeof cell === "string") {
                result.push(cell.trim().toUpperCase() === "TRUE");
            } else if (typeof cell === "number") {
                result.push(cell !== 0);
            }
        }
    }

    return result;
}

function getFreezeDateIso(workbook: ExcelScript.Workbook): string {
    const namedItem = workbook.getNamedItem("FreezeDate");
    const formula = namedItem.getFormula();

    const match = formula.match(/^=([^!]+)!\$?([A-Z]+)\$?(\d+)$/i);
    if (match) {
        const sheetName = match[1].replace(/'/g, "");
        const cellAddress = `${match[2]}${match[3]}`;
        const sheet = workbook.getWorksheet(sheetName);
        const raw = sheet.getRange(cellAddress).getValue() as CellValue;
        return normalizeDateCell(raw);
    }

    const fallback = namedItem.getValue() as CellValue;
    return normalizeDateCell(fallback);
}

function readMatrix(
    values: CellValue[][],
    blockRowIndex: number,
    blockColIndex: number,
    rowCount: number,
    colCount: number
): CellValue[][] {
    const matrix: CellValue[][] = [];

    for (let r = 0; r < rowCount; r++) {
        const outRow: CellValue[] = [];

        for (let c = 0; c < colCount; c++) {
            const rr = blockRowIndex + DATA_ROW_OFFSET + r;
            const cc = blockColIndex + DATA_COL_OFFSET + c;

            if (rr >= values.length || cc >= values[rr].length) {
                outRow.push(null);
            } else {
                outRow.push(normalizeCell(values[rr][cc]));
            }
        }

        matrix.push(outRow);
    }

    return matrix;
}

function buildHistoryExport(
    matrices: Record<string, CellValue[][]>,
    players: string[],
    dateCount: number
): Record<string, Record<string, CellValue[]>> {
    const history: Record<string, Record<string, CellValue[]>> = {};

    for (const metricName of Object.keys(matrices)) {
        const matrix = matrices[metricName];
        const metricHistory: Record<string, CellValue[]> = {};
        const rowCount = Math.min(players.length, matrix.length);

        for (let p = 0; p < rowCount; p++) {
            metricHistory[players[p]] = matrix[p]
                .slice(0, dateCount)
                .map((cell) => normalizeCell(cell));
        }

        history[metricName] = metricHistory;
    }

    return history;
}

function buildAllTables(
    matrices: Record<string, CellValue[][]>,
    players: string[],
    activeFlags: boolean[],
    dateCount: number
): Record<string, TableRow[]> {
    const tables: Record<string, TableRow[]> = {};

    for (const metricName of TABLE_METRICS) {
        const metricMatrix = matrices[metricName];
        if (!metricMatrix) {
            throw new Error(`Missing history block for ${metricName}`);
        }

        tables[metricName] = buildMetricTable(
            metricName,
            metricMatrix,
            matrices["RPHist"],
            matrices["TodayScoreHist"],
            matrices["CurrentStreakHist"],
            players,
            activeFlags,
            dateCount
        );
    }

    return tables;
}

function buildMetricTable(
    metricName: string,
    metricMatrix: CellValue[][],
    rpMatrix: CellValue[][],
    todayMatrix: CellValue[][],
    streakMatrix: CellValue[][],
    players: string[],
    activeFlags: boolean[],
    dateCount: number
): TableRow[] {
    const viewCol = dateCount - 1;
    const prevCol = viewCol - 1;

    const currentValues = getColumn(metricMatrix, viewCol);
    const previousValues = prevCol >= 0 ? getColumn(metricMatrix, prevCol) : [];

    const currentRanks = buildRanks(currentValues);
    const previousRanks = prevCol >= 0 ? buildRanks(previousValues) : [];

    const rpValues = getColumn(rpMatrix, viewCol);
    const todayValues = getColumn(todayMatrix, viewCol);
    const prevStreakValues = prevCol >= 0 ? getColumn(streakMatrix, prevCol) : [];

    const rowCount = Math.min(players.length, activeFlags.length, currentValues.length);
    const rows: (TableRow & { order: number })[] = [];

    for (let p = 0; p < rowCount; p++) {
        if (!activeFlags[p]) {
            continue;
        }

        const currentValue = currentValues[p];
        if (!hasValue(currentValue)) {
            continue;
        }

        const currentRank = currentRanks[p];
        const previousRank = prevCol >= 0 ? previousRanks[p] : null;

        rows.push({
            order: p,
            rank: currentRank,
            previousRank,
            change: buildChangeDisplay(previousRank, currentRank),
            player: players[p],
            streak: buildDispStreak(prevStreakValues[p], todayValues[p]),
            roundsPlayed: normalizeCell(rpValues[p]),
            value: normalizeCell(currentValue),
            form: buildLastFive(todayMatrix, p, viewCol),
        });
    }

    rows.sort((a, b) => {
        const diff = comparableNumber(b.value) - comparableNumber(a.value);
        if (diff !== 0) {
            return diff;
        }
        return a.order - b.order;
    });

    return rows.map(({ order, ...row }) => row);
}

function buildRanks(values: CellValue[]): number[] {
    const ranked = values.map((value, index) => ({
        value: comparableNumber(value),
        index,
    }));

    ranked.sort((a, b) => {
        const diff = b.value - a.value;
        if (diff !== 0) {
            return diff;
        }
        return a.index - b.index;
    });

    const ranks: number[] = new Array(values.length);
    for (let pos = 0; pos < ranked.length; pos++) {
        ranks[ranked[pos].index] = pos + 1;
    }

    return ranks;
}

function buildChangeDisplay(previousRank: number | null, currentRank: number): string {
    if (previousRank === null) {
        return "";
    }

    const delta = previousRank - currentRank;
    if (delta === 0) {
        return "";
    }

    return delta > 0 ? `▲${delta}` : `▼${Math.abs(delta)}`;
}

function buildDispStreak(prevStreakValue: CellValue, currentScoreValue: CellValue): string {
    const prevStreak = hasValue(prevStreakValue) ? Number(prevStreakValue) : 0;

    if (hasValue(currentScoreValue)) {
        return `🔥${prevStreak + 1}`;
    }

    if (prevStreak === 0) {
        return "";
    }

    return `⌛${prevStreak}`;
}

function buildLastFive(
    todayMatrix: CellValue[][],
    playerIndex: number,
    viewCol: number
): CellValue[] {
    const row = todayMatrix[playerIndex] ?? [];
    const result: CellValue[] = [];

    for (let offset = -4; offset <= 0; offset++) {
        const col = viewCol + offset;

        if (col < 0) {
            result.push(0);
            continue;
        }

        if (col >= row.length) {
            result.push(null);
            continue;
        }

        result.push(normalizeCell(row[col]));
    }

    return result;
}

function getColumn(matrix: CellValue[][], colIndex: number): CellValue[] {
    const column: CellValue[] = [];

    for (let r = 0; r < matrix.length; r++) {
        const row = matrix[r];
        column.push(colIndex < row.length ? normalizeCell(row[colIndex]) : null);
    }

    return column;
}

function normalizeCell(value: CellValue): CellValue {
    if (value === undefined || value === null || value === "") {
        return null;
    }
    return value;
}

function normalizeTextCell(value: CellValue): string {
    if (value === undefined || value === null || value === "") {
        return "";
    }
    return String(value).trim();
}

function normalizeDateCell(value: CellValue): string {
    if (value === undefined || value === null || value === "") {
        return "";
    }

    if (typeof value === "number") {
        return excelSerialToIsoDate(value);
    }

    const text = String(value).trim();
    if (text === "") {
        return "";
    }

    const parsed = Date.parse(text);
    if (!Number.isNaN(parsed)) {
        return new Date(parsed).toISOString().slice(0, 10);
    }

    return text;
}

function excelSerialToIsoDate(serial: number): string {
    const ms = Math.round((serial - 25569) * 86400 * 1000);
    return new Date(ms).toISOString().slice(0, 10);
}

function hasValue(value: CellValue): boolean {
    return !(value === null || value === undefined || value === "");
}

function comparableNumber(value: CellValue): number {
    if (value === null || value === undefined || value === "") {
        return Number.NEGATIVE_INFINITY;
    }

    if (typeof value === "number") {
        return value;
    }

    if (typeof value === "boolean") {
        return value ? 1 : 0;
    }

    const parsed = Number(value);
    return Number.isNaN(parsed) ? Number.NEGATIVE_INFINITY : parsed;
}