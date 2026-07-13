export async function loadLeague() {
    const response = await fetch("./data/league.json", {
        cache: "no-store",
    });

    if (!response.ok) {
        throw new Error(`Failed to load league.json (${response.status})`);
    }

    return response.json();
}