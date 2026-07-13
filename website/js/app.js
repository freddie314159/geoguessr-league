import {loadLeague} from "./api.js";
import {renderMetricTable} from "./table.js";

function setError(message) {
    const error = document.getElementById("error");
    error.hidden = !message;
    error.textContent = message || "";
}

function populateMetricSelect(select, data) {
    const order = Array.isArray(data.metricOrder) ? data.metricOrder : [];
    const labels = data.metricLabels || {};

    select.innerHTML = order.map((metric) => {
        const label = labels[metric] ?? metric;
        return `<option value="${metric}">${label}</option>`;
    })
    .join("");
}

async function main() {
    const meta = document.getElementById("meta");
    const metricSelect = document.getElementById("metric");
    const tableContainer = document.getElementById("table");

    try {
        const data = await loadLeague();

        meta.textContent = `Last updated: ${data.freezeDate || "unknown"}`;

        populateMetricSelect(metricSelect, data);

        const defaultMetric = 
            data.metricOrder?.includes("TotScoreHist")
                ? "TotScoreHist"
                : data.metricOrder?.[0];

        if (defaultMetric) {
            metricSelect.value = defaultMetric;
        }

        const render = () => {
            renderMetricTable(tableContainer, data, metricSelect.value);
        };

        metricSelect.addEventListener("change", render);
        render();
    } catch (err) {
        setError(err instanceof Error ? err.message : String(err));
    }
}

document.addEventListener("DOMContentLoaded", main);