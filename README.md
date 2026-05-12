# LinkedIn Automation Agent

LinkedIn থেকে company এবং contact data automatically collect করার unified pipeline।

---

## Pipeline Overview (কীভাবে কাজ করে)

```
STEP 1  →  LinkedIn search থেকে Company URL collect করে
STEP 2  →  প্রতিটি Company এর About + People page scrape করে
STEP 3  →  Google / Bing search করে individual Contact LinkedIn URL খোঁজে
STEP 4  →  Duplicate remove করে final output তৈরি করে
STEP 5  →  প্রতিটি Contact profile visit করে Company Name match verify করে
```

---

## Files (কোন file কী কাজ করে)

| File | কাজ |
|------|-----|
| `linkedin_agent.py` | **Main script** — সব step এক জায়গায় |
| `linkedin.txt` | LinkedIn email + password (credentials) |
| `agent_state.json` | Checkpoint — কোন step এ আছে track করে (auto-generated) |
| `agent_companies.xlsx` | Step 1 output — Company URLs |
| `agent_people.xlsx` | Step 2 output — Company info + contacts |
| `agent_profiles.xlsx` | Step 3 output — Contact LinkedIn URLs |
| `agent_final.xlsx` | Step 4 output — Deduplicated final data |
| `agent_verified.xlsx` | Step 5 output — Company match verified (Yes/No) |
| `Old/` | পুরনো individual scripts (backup) |

---

## Setup (প্রথমবার চালানোর আগে)

### 1. Python packages install করুন

```powershell
pip install selenium webdriver-manager pandas openpyxl
```

### 2. linkedin.txt চেক করুন

File এর **প্রথম দুই line** হতে হবে:
```
your_email@gmail.com
your_password
```

> **Note:** Password এর পর যা আছে সেটা notes — script শুধু প্রথম দুই line পড়ে।

### 3. linkedin_agent.py এর CONFIG section ঠিক করুন

File এর উপরের দিকে CONFIG section আছে — এগুলো নিজের মতো পরিবর্তন করুন:

```python
# কোন LinkedIn search URL থেকে company collect করবেন
LINKEDIN_SEARCH_URL = "https://www.linkedin.com/search/results/companies/?..."

# কত page scrape করবেন (প্রতি page এ ~10 company)
TOTAL_SEARCH_PAGES = 5

# Director-level হিসেবে কোন designation count হবে
DIRECTOR_KEYWORDS = ["ceo", "director", "vp", "manager", ...]
```

---

## Run করুন

PowerShell বা Command Prompt খুলে:

```powershell
cd "d:\Backup-11-26-25\Linkedin Automation"
python linkedin_agent.py
```

### প্রথমবার চললে কী হবে:

1. Chrome automatically open হবে
2. LinkedIn login হবে automatically (`linkedin.txt` থেকে credentials নিয়ে)
3. যদি LinkedIn verification/CAPTCHA আসে → terminal এ message আসবে → Chrome এ solve করুন → Enter চাপুন
4. বাকি সব automatically চলবে

---

## Resume (মাঝে বন্ধ হয়ে গেলে)

Script crash বা বন্ধ হলে **আবার same command দিন** — সে আগের জায়গা থেকে শুরু করবে:

```powershell
python linkedin_agent.py
```

`agent_state.json` file দেখে কোন step এ ছিল বুঝে নেয়। নতুন করে শুরু করতে চাইলে এই file delete করুন।

---

## Output Columns (Final Files এ কী কী থাকে)

### agent_final.xlsx / agent_verified.xlsx

| Column | মানে |
|--------|------|
| `LinkedIn URL` | Company এর LinkedIn URL |
| `Company Name` | Company এর নাম |
| `Website` | Company website |
| `Company Phone` | Phone number |
| `Industry` | Industry type |
| `Number of Employees` | Employee count |
| `Street Address` | ঠিকানা |
| `City`, `State`, `Zip Code`, `Country` | Location details |
| `Name` | Contact এর নাম |
| `Designation` | Contact এর designation |
| `Director Level` | Yes/No — Director/VP/CEO level কিনা |
| `Contact LinkedIn URL` | Contact এর LinkedIn profile URL |
| `Matched Title` | Google/Bing search থেকে পাওয়া title |
| `Search Source` | Google বা Bing |
| `First Name` | Profile থেকে extracted first name (Step 5) |
| `Last Name` | Profile থেকে extracted last name (Step 5) |
| `Profile Title` | Profile এর current job title (Step 5) |
| `Company Match` | **Yes/No** — Company name profile এ আছে কিনা (Step 5) |

---

## Google CAPTCHA হলে কী করবেন

- Script নিজেই detect করে
- **৩ বার** CAPTCHA আসলে automatically **Bing** এ switch করে
- Manual solve করতে চাইলে: Chrome এ solve করুন → terminal এ Enter চাপুন

---

## Director Level Filter

এই designations automatically `Director Level = Yes` হয়:

`CEO, COO, CFO, CTO, CMO, Founder, President, Owner, Partner, Director, VP, Vice President, Manager, Head of, Chief, Executive, Superintendent, Administrator`

নতুন keyword যোগ করতে `linkedin_agent.py` এর `DIRECTOR_KEYWORDS` list এ add করুন।

---

## Common Problems

| সমস্যা | সমাধান |
|--------|--------|
| `ModuleNotFoundError: selenium` | `pip install selenium webdriver-manager` চালান |
| LinkedIn login হচ্ছে না | `linkedin.txt` এর প্রথম দুই line check করুন |
| Script হঠাৎ বন্ধ হয়ে গেছে | আবার `python linkedin_agent.py` চালান — resume হবে |
| নতুন করে শুরু করতে চাই | `agent_state.json` delete করুন, তারপর চালান |
| Step 2 এ data নেই | `agent_companies.xlsx` check করুন — company URL আছে কিনা |
| Chrome version mismatch | `pip install --upgrade webdriver-manager` চালান |

---

## Folder Structure

```
Linkedin Automation/
├── linkedin_agent.py          ← Main script (এটাই চালান)
├── linkedin.txt               ← Credentials (email + password)
├── README.md                  ← এই file
│
├── agent_state.json           ← Auto-generated checkpoint
├── agent_companies.xlsx       ← Step 1 output
├── agent_people.xlsx          ← Step 2 output
├── agent_profiles.xlsx        ← Step 3 output
├── agent_final.xlsx           ← Step 4 output (deduplicated)
├── agent_verified.xlsx        ← Step 5 output (verified)
│
└── Old/                       ← পুরনো individual scripts (backup)
    ├── step1.py
    ├── pet_automation.py
    ├── pet_automation2.py
    ├── step3.py
    └── duplicate.py
```
