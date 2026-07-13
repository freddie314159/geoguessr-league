function escapeText(value) {
    return String(value ?? "");
}

function changeClass(change) {
    if (!change) return "change-none";
    if (change.startsWith("▲")) return "change-up";
    if (change.startsWith("▼")) return "change-down";
    return "change-none";
}

export function renderMetricTable(container, data, metricName) {
    const rows = data.tables?.[metricName] ?? [];
    const metricLabel = data.metricLabels?.[metricName] ?? metricName;

    const wrap = document.createElement("div");
    wrap.className = "table-wrap";

    const title = document.createElement("h2");
    title.textContent = metricLabel;
    title.style.margin = "0 0 12px";
    wrap.appendChild(title);

    if (rows.length === 0) {
        const empty = document.createElement("p");
        empty.textContent = "No rows available for this metric.";
        wrap.appendChild(empty);
        container.replaceChildren(wrap);
        return;
    }

    const table = document.createElement("table");

    table.innerHTML = `
        <thead>
            <tr>
                <th class="rank">Rank</th>
                <th>Change</th>
                <th>Player</th>
                <th>Streak</th>
                <th>Rounds</th>
                <th>Value</th>
                <th>Last 5</th>
            </tr>
        </thead>
    `;

    const tbody = document.createElement("tbody");
    
    rows.forEach((row) => {
        const tr = document.createElement("tr");

        const form = Array.isArray(row.form) ? row.form : [];
        const formHtml = form
            .map((value) => `<span class="form-box ${value == null ? "empty" : "filled"}"></span>`)
            .join("");

        tr.innerHTML = `
            <td class="rank">${escapeText(row.rank)}</td>
            <td class="${changeClass(row.change)}">${escapeText(row.change)}</td>
            <td class="player">${escapeText(row.player)}</td>
            <td class="streak">${escapeText(row.streak)}</td>
            <td>${escapeText(row.roundsPlayed)}</td>
            <td>${escapeText(row.value)}</td>
            <td><span class="form">${formHtml}</span></td>
        `;

        tbody.appendChild(tr);
    });
    
    table.appendChild(tbody);
    wrap.appendChild(table);
    container.replaceChildren(wrap);
}