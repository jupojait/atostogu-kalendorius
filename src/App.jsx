import React, { useEffect, useMemo, useState } from "react";
import { useMsal } from "@azure/msal-react";
import "./App.css";

const MONTHS = [
  "Sausis", "Vasaris", "Kovas", "Balandis",
  "Gegužė", "Birželis", "Liepa", "Rugpjūtis",
  "Rugsėjis", "Spalis", "Lapkritis", "Gruodis"
];

const WEEK_DAYS = ["Pr", "An", "Tr", "Kt", "Pn", "Št", "Sk"];

const COLORS = [
  "#22c55e", "#3b82f6", "#f97316", "#8b5cf6",
  "#ec4899", "#14b8a6", "#eab308", "#ef4444"
];

function isoDate(year, month, day) {
  return `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function initials(name = "") {
  return name
    .split(" ")
    .filter(Boolean)
    .slice(0, 2)
    .map(x => x[0])
    .join("")
    .toUpperCase();
}

function getMonthDays(year, month) {
  const first = new Date(year, month, 1);
  const count = new Date(year, month + 1, 0).getDate();

  let offset = first.getDay();
  offset = offset === 0 ? 6 : offset - 1;

  const days = Array(offset).fill(null);

  for (let d = 1; d <= count; d++) {
    const date = new Date(year, month, d);
    days.push({
      day: d,
      iso: isoDate(year, month, d),
      weekend: date.getDay() === 0 || date.getDay() === 6
    });
  }

  return days;
}

function dayBackground(people) {
  if (people.length === 0) return {};

if (people.length === 1) {
  return {
    background: people[0].color
  };
}

  if (people.length === 2) {
    return {
      background: `linear-gradient(135deg, ${people[0].color} 0 50%, ${people[1].color} 50% 100%)`
    };
  }

  if (people.length === 3) {
    return {
      background: `conic-gradient(${people[0].color} 0deg 120deg, ${people[1].color} 120deg 240deg, ${people[2].color} 240deg 360deg)`
    };
  }

  if (people.length === 4) {
    return {
      background: `conic-gradient(${people[0].color} 0deg 90deg, ${people[1].color} 90deg 180deg, ${people[2].color} 180deg 270deg, ${people[3].color} 270deg 360deg)`
    };
  }

  return {
    background: `repeating-linear-gradient(45deg, ${people[0].color} 0 6px, ${people[1].color} 6px 12px, ${people[2].color} 12px 18px, ${people[3].color} 18px 24px)`
  };
}

export default function App() {
  const { instance, accounts } = useMsal();

  const user = accounts[0];

  const currentYear = new Date().getFullYear();

  const [year, setYear] = useState(currentYear);
  const [loading, setLoading] = useState(false);
  const [officeUsers, setOfficeUsers] = useState([]);
  const [selectedPersonId, setSelectedPersonId] = useState("");
  const [vacations, setVacations] = useState({});
  const [currentUserSettings, setCurrentUserSettings] = useState(null);
  const [activePage, setActivePage] = useState("calendar");
const [settingsRows, setSettingsRows] = useState([]);
const [newUserEmail, setNewUserEmail] = useState("");
const [newSettingsUserId, setNewSettingsUserId] = useState("");

  async function login() {
    await instance.loginRedirect({
      scopes: ["User.Read", "User.ReadBasic.All"]
    });
  }

  async function logout() {
    await instance.logoutRedirect();
  }

  useEffect(() => {

    async function loadUsers() {
  if (!user) return;

  try {
    setLoading(true);

    const token = await instance.acquireTokenSilent({
      scopes: ["User.Read", "User.ReadBasic.All"],
      account: user
    });

    const response = await fetch(
      "https://graph.microsoft.com/v1.0/users?$top=50&$select=displayName,mail,userPrincipalName,userType,jobTitle,department,accountEnabled",
      {
        headers: {
          Authorization: `Bearer ${token.accessToken}`
        }
      }
    );

    const data = await response.json();
console.log("GRAPH FULL DATA:", data);


    console.log("GRAPH RESPONSE:", data);

    if (data.error) {
      alert(data.error.message);
      return;
    }

const mappedUsers = (data.value || [])
  .filter(u => u.displayName)
  .filter(u => !u.userPrincipalName?.includes("#EXT#"))
  .filter(u => u.department)
  .map((u, index) => ({
    id: u.id || u.userPrincipalName,
    name: u.displayName,
    
    email: u.mail || u.userPrincipalName,
    initials: initials(u.displayName),
    color: COLORS[index % COLORS.length]
  }));

    setOfficeUsers(mappedUsers);

  const currentPerson = mappedUsers.find(
  x => x.email?.toLowerCase() === user.username?.toLowerCase()
);

if (currentPerson) {
  setSelectedPersonId(currentPerson.id);
} else if (mappedUsers.length > 0) {
  setSelectedPersonId(mappedUsers[0].id);
}

  } catch (e) {
    console.error("LOAD USERS ERROR:", e);
    alert(e.message);
  } finally {
    setLoading(false);
  }
}

    loadUsers();
  }, [user, instance]);

  useEffect(() => {
  if (officeUsers.length > 0) {
    loadVacationsFromSharePoint();
  }
}, [officeUsers]);
useEffect(() => {
  if (user) {
    loadUserSettings();
  }
}, [user]);

  const selectedPerson = officeUsers.find(x => x.id === selectedPersonId);

  const selectedCount = selectedPersonId
    ? (vacations[selectedPersonId] || []).length
    : 0;

  function peopleOnDate(date) {
    return officeUsers.filter(person => (vacations[person.id] || []).includes(date));
  }
const SITE_ID =
"jupoja.sharepoint.com,e14a2f73-54d1-423b-aaab-a450444fd9f4,36013266-dbed-4f84-9400-e34c929c4207";

const LIST_ID =
"28704355-4855-4053-b927-4893c877a41c";
const SETTINGS_LIST_ID =
"959c37e5-6500-4cb3-b06d-884013ad2d1a";
async function saveVacation(date) {

  const token = await instance.acquireTokenSilent({
    scopes: [
      "User.Read",
      "Sites.ReadWrite.All"
    ],
    account: user
  });

  const selectedPerson = officeUsers.find(
    x => x.id === selectedPersonId
  );

  await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token.accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
fields: {
  Title: selectedPerson.name,
  Metai: year,
  AtostoguData: date,
  DarbuotojoEmail: selectedPerson.email
}
      })
    }
  );

}

async function deleteVacation(date) {

  const token = await instance.acquireTokenSilent({
    scopes: [
      "User.Read",
      "Sites.ReadWrite.All"
    ],
    account: user
  });

  const selectedPerson = officeUsers.find(
    x => x.id === selectedPersonId
  );

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields`,
    {
      headers: {
        Authorization: `Bearer ${token.accessToken}`
      }
    }
  );

  const data = await response.json();

  const item = (data.value || []).find(x => {
    const f = x.fields;

    return (
      f.DarbuotojoEmail?.toLowerCase() === selectedPerson.email?.toLowerCase()
      &&
      f.AtostoguData?.slice(0, 10) === date
    );
  });

  if (!item) return;

  await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items/${item.id}`,
    {
      method: "DELETE",
      headers: {
        Authorization: `Bearer ${token.accessToken}`
      }
    }
  );
}


async function loadVacationsFromSharePoint() {
  const token = await instance.acquireTokenSilent({
    scopes: ["User.Read", "Sites.ReadWrite.All"],
    account: user
  });

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields($select=Title,AtostoguData,DarbuotojoEmail)&$top=999`,
    {
      headers: {
        Authorization: `Bearer ${token.accessToken}`
      }
    }
  );

  const data = await response.json();

  const loaded = {};

  (data.value || []).forEach(item => {
    const f = item.fields;

    const person = officeUsers.find(
      u => u.email?.toLowerCase() === f.DarbuotojoEmail?.toLowerCase()
    );

    if (!person) return;

    if (!loaded[person.id]) {
      loaded[person.id] = [];
    }

    loaded[person.id].push(f.AtostoguData?.slice(0, 10));
  });

  setVacations(loaded);
}
async function loadUserSettings() {
  const token = await instance.acquireTokenSilent({
    scopes: ["User.Read", "Sites.ReadWrite.All"],
    account: user
  });

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${SETTINGS_LIST_ID}/items?expand=fields`,
    {
      headers: {
        Authorization: `Bearer ${token.accessToken}`
      }
    }
  );

  const data = await response.json();

  const settings = (data.value || []).find(item =>
    item.fields.UserEmail?.toLowerCase() === user.username?.toLowerCase()
  );

  setCurrentUserSettings(settings?.fields || null);
}
async function loadAllSettings() {
  const token = await instance.acquireTokenSilent({
    scopes: ["User.Read", "Sites.ReadWrite.All"],
    account: user
  });

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${SETTINGS_LIST_ID}/items?expand=fields&$top=999`,
    {
      headers: {
        Authorization: `Bearer ${token.accessToken}`
      }
    }
  );

  const data = await response.json();

  setSettingsRows(
    (data.value || []).map(item => ({
      itemId: item.id,
      title: item.fields.Title,
      userEmail: item.fields.UserEmail,
      canManageOthers: item.fields.CanManageOthers === true,
      isAdmin: item.fields.IsAdmin === true
    }))
  );
}
async function toggleDate(date) {
  if (!selectedPersonId) return;
  const selectedPerson = officeUsers.find(
  x => x.id === selectedPersonId
);

if (
  !canManageOthers &&
  selectedPerson?.email?.toLowerCase() !== user.username?.toLowerCase()
) {
  alert("Galite žymėti tik savo atostogas.");
  return;
}

  const existing = vacations[selectedPersonId] || [];
  const exists = existing.includes(date);

  if (exists) {
    await deleteVacation(date);
  } else {
    await saveVacation(date);
  }

  setVacations(prev => {
    const current = prev[selectedPersonId] || [];

    return {
      ...prev,
      [selectedPersonId]: exists
        ? current.filter(d => d !== date)
        : [...current, date].sort()
    };
  });
}
async function saveUserSetting(row) {
  const token = await instance.acquireTokenSilent({
    scopes: ["User.Read", "Sites.ReadWrite.All"],
    account: user
  });

  await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${SETTINGS_LIST_ID}/items/${row.itemId}/fields`,
    {
      method: "PATCH",
      headers: {
        Authorization: `Bearer ${token.accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        CanManageOthers: row.canManageOthers,
        IsAdmin: row.isAdmin
      })
    }
  );

  await loadUserSettings();
  await loadAllSettings();
}

async function createUserSetting() {
  if (!newSettingsUserId) return;

  const person = officeUsers.find(x => x.id === newSettingsUserId);
  if (!person) return;

  const token = await instance.acquireTokenSilent({
    scopes: ["User.Read", "Sites.ReadWrite.All"],
    account: user
  });

  await fetch(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${SETTINGS_LIST_ID}/items`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token.accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        fields: {
          Title: person.name,
          UserEmail: person.email,
          CanManageOthers: false,
          IsAdmin: false
        }
      })
    }
  );

  setNewSettingsUserId("");
  await loadAllSettings();
}



const isAdmin =
  currentUserSettings?.IsAdmin === true;

const canManageOthers =
  currentUserSettings?.CanManageOthers === true;
useEffect(() => {
  if (isAdmin && user) {
    loadAllSettings();
  }
}, [isAdmin, user]);

  const totalMarkedDays = useMemo(() => {
    return Object.values(vacations).reduce((sum, arr) => sum + arr.length, 0);
  }, [vacations]);

useEffect(() => {
  if (!user) {
    instance.loginRedirect({
      scopes: ["User.Read", "User.ReadBasic.All"]
    });
  }
}, [user, instance]);

if (!user) {
  return (
    <div className="loginPage">
      <div className="loginCard">
        <h1>Kraunama...</h1>
      </div>
    </div>
  );
}

  return (
    <div className="appShell">
      <aside className="sidebar">
        <div className="brand">
          <div className="brandIcon">A</div>
          <div>
            <strong>Atostogos</strong>
            <span>Team planner</span>
          </div>
        </div>

        <nav>
          <div
  className={activePage === "calendar" ? "navItem active" : "navItem"}
  onClick={() => setActivePage("calendar")}
>
  Kalendorius
</div>

<div className="navItem">Mano dienos</div>
<div className="navItem">Komanda</div>

{isAdmin && (
  <div
    className={activePage === "settings" ? "navItem active" : "navItem"}
    onClick={() => setActivePage("settings")}
  >
    Nustatymai
  </div>
)}
        </nav>

        <div className="sidebarFooter">
          <strong>{user.name}</strong>
          <span>{user.username}</span>
          <button onClick={logout}>Atsijungti</button>
        </div>
      </aside>

      <main className="main">
        {activePage === "settings" ? (
  <div>
    <header className="topbar">
      <div>
        <h1>Nustatymai</h1>
        <p>Kas gali žymėti atostogas už kitus.</p>
      </div>
    </header>
<div className="settingsAddBar">
 
<select
  value={newSettingsUserId}
  onChange={e => setNewSettingsUserId(e.target.value)}
>
  <option value="">Pasirink vartotoją</option>

  {officeUsers
    .filter(person =>
      !settingsRows.some(row =>
        row.userEmail?.toLowerCase() === person.email?.toLowerCase()
      )
    )
    .map(person => (
      <option key={person.id} value={person.id}>
        {person.name} — {person.email}
      </option>
    ))}
</select>

<button onClick={createUserSetting}>
  Pridėti
</button>

</div>
    <div className="settingsTable">
      <div className="settingsHeader">
        <div>Vartotojas</div>
        <div>Email</div>
        <div>Už kitus</div>
        <div>Admin</div>
        <div></div>
      </div>

      {settingsRows.map(row => (
        <div className="settingsRow" key={row.itemId}>
          <div>{row.title}</div>
          <div>{row.userEmail}</div>

          <input
            type="checkbox"
            checked={row.canManageOthers}
            onChange={e =>
              setSettingsRows(prev =>
                prev.map(x =>
                  x.itemId === row.itemId
                    ? { ...x, canManageOthers: e.target.checked }
                    : x
                )
              )
            }
          />

          <input
            type="checkbox"
            checked={row.isAdmin}
            onChange={e =>
              setSettingsRows(prev =>
                prev.map(x =>
                  x.itemId === row.itemId
                    ? { ...x, isAdmin: e.target.checked }
                    : x
                )
              )
            }
          />

          <button onClick={() => saveUserSetting(row)}>
            Išsaugoti
          </button>
        </div>
      ))}
    </div>
  </div>
) : (
  <>
        <header className="topbar">
          <div>
            <h1>Metinis atostogų kalendorius</h1>
            <p>Žymėk atostogas ir iš karto matyk persidengimus komandoje.</p>
          </div>

          <div className="yearPicker">
            <button onClick={() => setYear(year - 1)}>‹</button>
            <strong>{year}</strong>
            <button onClick={() => setYear(year + 1)}>›</button>
          </div>
        </header>

        <section className="summaryGrid">
          <div className="summaryCard">
            <span>Prisijungęs</span>
            <strong>{user.name}</strong>
            <small>{user.username}</small>
          </div>

          <div className="summaryCard">
            <span>Žymi kam</span>
{canManageOthers ? (
  <select
    value={selectedPersonId}
    onChange={e => setSelectedPersonId(e.target.value)}
  >
    {officeUsers.map(person => (
      <option key={person.id} value={person.id}>
        {person.name}
      </option>
    ))}
  </select>
) : (
  <strong>{user.name}</strong>
)}
          </div>

          <div className="summaryCard">
            <span>Pasirinkta šiam asmeniui</span>
            <strong>{selectedCount}</strong>
            <small>dienos</small>
          </div>

          <div className="summaryCard">
            <span>Viso pažymėta</span>
            <strong>{totalMarkedDays}</strong>
            <small>dienos komandoje</small>
          </div>
        </section>

        {loading && <div className="notice">Kraunami Office 365 vartotojai...</div>}

        <section className="peopleLegend">
          {officeUsers.map(person => (
            <button
              key={person.id}
              className={person.id === selectedPersonId ? "personChip selected" : "personChip"}
              onClick={() => setSelectedPersonId(person.id)}
            >
              <span style={{ backgroundColor: person.color }}>{person.initials}</span>
              <div>
                <strong>{person.name}</strong>
                <small>{person.email}</small>
              </div>
            </button>
          ))}
        </section>

        <section className="calendarGrid">
          {MONTHS.map((month, monthIndex) => (
            <div className="monthCard" key={month}>
              <div className="monthHeader">
                <h3>{month}</h3>
              </div>

              <div className="weekdays">
                {WEEK_DAYS.map(day => (
                  <div key={day}>{day}</div>
                ))}
              </div>

              <div className="days">
                {getMonthDays(year, monthIndex).map((date, index) => {
                  if (!date) return <div key={index} />;

                  const people = peopleOnDate(date.iso);
                  const selectedByCurrent = people.some(p => p.id === selectedPersonId);

                  return (
                    <button
                      key={date.iso}
                      className={[
                        "day",
                        date.weekend ? "weekend" : "",
                        people.length ? "hasVacation" : "",
                        selectedByCurrent ? "selectedByCurrent" : ""
                      ].join(" ")}
                      onClick={() => toggleDate(date.iso)}
                      title={people.length ? people.map(p => p.name).join(", ") : date.iso}
                    >
                      {people.length > 0 && (
                        <span className="dayFill" style={dayBackground(people)} />
                      )}

                      <span className="dayNumber">{date.day}</span>

                      {people.length > 4 && <span className="moreBadge">4+</span>}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
        </section>

        <section className="overlapLegend">
          <strong>Persidengimų žymėjimas</strong>
          <span className="sample one"></span> 1 žmogus
          <span className="sample two"></span> 2 žmonės
          <span className="sample three"></span> 3 žmonės
          <span className="sample four"></span> 4 žmonės
          <span className="sample many"></span> daugiau nei 4
        </section>
          </>
)}
      </main>
    </div>
  );
}