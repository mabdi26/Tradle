import React, { useMemo, useState } from "react";
import { ResponsiveContainer, Treemap, Tooltip } from "recharts";

// ================= TYPES =================

type ExportCategory = { name: string; value: number };

type Country = {
  name: string;
  iso: string;
  continent: string;
  lat: number;
  lon: number;
  exports: ExportCategory[];
};

type Guess = {
  country: Country;
  distance: number;
  similarity: number;
};

// ================= CONSTANTS =================

const MAX_GUESSES = 6;

const COLORS = [
  "#FFD43B",
  "#FFA63F",
  "#FF7D05",
  "#6366F1",
  "#22C55E",
  "#06B6D4"
];

// ================= DATA =================

const COUNTRIES: Country[] = [
  {
    name: "United States",
    iso: "USA",
    continent: "North America",
    lat: 37,
    lon: -95,
    exports: [
      { name: "Technology", value: 30 },
      { name: "Machinery", value: 20 },
      { name: "Finance", value: 15 },
      { name: "Healthcare", value: 10 }
    ]
  },
  {
    name: "China",
    iso: "CHN",
    continent: "Asia",
    lat: 35,
    lon: 104,
    exports: [
      { name: "Electronics", value: 40 },
      { name: "Manufacturing", value: 25 },
      { name: "Textiles", value: 10 }
    ]
  },
  {
    name: "Germany",
    iso: "DEU",
    continent: "Europe",
    lat: 51,
    lon: 10,
    exports: [
      { name: "Automobiles", value: 35 },
      { name: "Engineering", value: 20 },
      { name: "Chemicals", value: 12 }
    ]
  }
];

// ================= UTILS =================

const distance = (a: Country, b: Country) =>
  Math.round(Math.sqrt((a.lat - b.lat) ** 2 + (a.lon - b.lon) ** 2) * 111);

function similarity(a: Country, b: Country) {
  let diff = 0;
  const cats = new Set([...a.exports, ...b.exports].map(e => e.name));

  cats.forEach(c => {
    const av = a.exports.find(e => e.name === c)?.value || 0;
    const bv = b.exports.find(e => e.name === c)?.value || 0;
    diff += Math.abs(av - bv);
  });

  return Math.max(0, Math.round(100 - diff));
}

const heat = (d: number) => {
  if (d < 500) return "🔥 Hot";
  if (d < 2000) return "🌡 Warm";
  if (d < 5000) return "❄ Cool";
  return "🧊 Cold";
};

const randomCountry = () => COUNTRIES[Math.floor(Math.random() * COUNTRIES.length)];

const findCountry = (name: string) =>
  COUNTRIES.find(c => c.name.toLowerCase() === name.toLowerCase());

// ================= TREEMAP =================

function ExportTreemap({ data }: { data: ExportCategory[] }) {
  const formatted = data.map((d, i) => ({ ...d, size: d.value, fill: COLORS[i % COLORS.length] }));

  return (
    <div className="rounded-2xl overflow-hidden border bg-gray-900">
      <ResponsiveContainer width="100%" height={350}>
        <Treemap data={formatted} dataKey="size">
          <Tooltip formatter={(v: any, n: any) => [`${v}%`, n]} />
        </Treemap>
      </ResponsiveContainer>
    </div>
  );
}

// ================= UI =================

function Header({ reset }: { reset: () => void }) {
  return (
    <div className="flex justify-between items-center border-b border-gray-800 pb-3">
      <h1 className="text-3xl font-bold bg-gradient-to-r from-yellow-300 to-orange-400 bg-clip-text text-transparent">
        Export Oracle
      </h1>
      <button onClick={reset} className="px-4 py-2 bg-gray-800 rounded-lg hover:bg-gray-700">
        New Game
      </button>
    </div>
  );
}

function GuessRow({ g }: { g: Guess }) {
  return (
    <div className="flex justify-between p-3 rounded-lg bg-black border border-gray-800 hover:scale-[1.01] transition">
      <span>{g.country.name}</span>
      <span className="text-sm opacity-80">
        {g.distance} km • {heat(g.distance)} • {g.similarity}%
      </span>
    </div>
  );
}

// ================= APP =================

export default function App() {
  const [target, setTarget] = useState(randomCountry);
  const [input, setInput] = useState("");
  const [guesses, setGuesses] = useState<Guess[]>([]);
  const [status, setStatus] = useState("playing");

  const options = useMemo(() => COUNTRIES.map(c => c.name), []);

  function reset() {
    setTarget(randomCountry());
    setGuesses([]);
    setStatus("playing");
  }

  function submit(e: React.FormEvent) {
    e.preventDefault();
    if (status !== "playing") return;

    const c = findCountry(input);
    if (!c) return;

    const g: Guess = {
      country: c,
      distance: distance(c, target),
      similarity: similarity(c, target)
    };

    const next = [...guesses, g];
    setGuesses(next);
    setInput("");

    if (c.iso === target.iso) setStatus("won");
    else if (next.length >= MAX_GUESSES) setStatus("lost");
  }

  async function share() {
    const rows = guesses.map(g => `${g.distance}km`).join("\n");

    const text = `Export Oracle ${status === "won" ? guesses.length : "X"}/${MAX_GUESSES}\n${rows}\nAnswer: ${target.name}`;

    await navigator.clipboard.writeText(text);
  }

  return (
    <div className="min-h-screen bg-[#0a0a0a] text-white p-6 flex justify-center">
      <div className="w-full max-w-6xl space-y-6">

        <Header reset={reset} />

        <div className="grid md:grid-cols-2 gap-6">

          <ExportTreemap data={target.exports} />

          <div className="bg-gray-900 border border-gray-800 rounded-2xl p-5 space-y-4">

            <form onSubmit={submit} className="flex gap-2">
              <input
                className="flex-1 p-3 rounded-lg bg-black border border-gray-700"
                value={input}
                list="countries"
                onChange={(e) => setInput(e.target.value)}
                placeholder="Guess a country..."
              />

              <datalist id="countries">
                {options.map(o => <option key={o} value={o} />)}
              </datalist>

              <button className="px-4 bg-orange-500 rounded-lg">Guess</button>
            </form>

            <button onClick={share} className="w-full bg-yellow-500 py-2 rounded-lg">Share</button>

            <div className="space-y-2">
              {guesses.map((g, i) => <GuessRow key={i} g={g} />)}
            </div>

            {status !== "playing" && (
              <div className="text-center pt-4 text-lg font-semibold">
                {status === "won" ? "🎉 Correct!" : `Answer: ${target.name}`}
              </div>
            )}

          </div>
        </div>
      </div>
    </div>
  );
}

/* INSTALL:
npm install recharts
*/