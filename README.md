# Mysore District Athletic Association (MDAA) — Athlete Entry Portal
## District Meet to Select Athletes for the 20th National Inter District Junior Athletics Meet 2026 (NIDJAM)
**Official Selection & Compliance Guidelines under AFI Circular No. 88/09-2026**

### Overview
Mysore District Athletic Association (MDAA) is conducting the official District Athletics Meet to select athletes to represent Mysore District at the **20th National Inter District Junior Athletics Meet 2026 (NIDJAM)**, held at **Woxsen University, Kamkole, Sadasivpet, Hyderabad, Telangana** from **14 – 16 November 2026**, organized by the Athletics Federation of India (AFI) in collaboration with the Government of Telangana.

---

### Critical Notice: AFI UID Policy
> **IMPORTANT NOTE TO ALL ATHLETES**
> **AFI UID IS COMPULSORY FOR NIDJAM**
> 1. Only athletes with AFI UID can participate in NIDJAM Meet.
> 2. Athletes without AFI UID can participate in this competition and if they win, they will get Certificate and Medal, but they will NOT be allowed to take part in NIDJAM Meet.
> 3. You can register for AFI UID at AFI Website: [www.indianathletics.in](https://www.indianathletics.in) — AFI UID Registration (or [ors.indianathletics.in](https://ors.indianathletics.in/#/login)).
> 4. All athletes are requested to register and get your AFI UID immediately before the meet.
> **No AFI UID = No NIDJAM**

---

### Meet Rules & Categories

1. **Age Brackets & Cut-off Dates (as on 16 November 2026)**
   - **Under 14 (U-14)**: Born between **17/11/2012** & **16/11/2014** (Minimum age 12 years and above as on 16 Nov 2026; boys and girls under 12 are strictly ineligible).
   - **Under 16 (U-16)**: Born between **17/11/2010** & **16/11/2012**.
   - Original Birth Certificate (issued by Municipal Corporation / Nagar Palika / District Office) or 10th Standard Certificate mandatory for Age Verification on **13 November 2026**.
   - Zero tolerance on over-aging: immediate disqualification.

2. **Event Rules & Restrictions**
   - **Under 14**: Must choose exactly **one Triathlon group** (Tri A, Tri B, or Tri C) + **Kids Javelin Throw** (with 5m runway, mandatory for all U-14 athletes):
     - *Tri A*: 60m, Long Jump (5m approach), High Jump (Scissor only)
     - *Tri B*: 60m, Long Jump (5m approach), Back Throw (1kg Shot Put)
     - *Tri C*: 60m, Long Jump (5m approach), 600m
     - *Kids Javelin Throw*: 5-meter runway (Auto-included)
   - **Under 16**: Can enter **EXACTLY ONE EVENT**:
     - *Boys*: 60m, 80m Hurdles (0.838m), 600m, High Jump (Scissor Only), Long Jump (5m Approach), 4Kg Shot Put (Standing), 600gms Javelin Throw (10m Approach), Pentathlon (60m · 80mH · LJ · SP · 600m).
     - *Girls*: 60m, 80m Hurdles (0.762m), 600m, High Jump (Scissor Only), Long Jump (5m Approach), 3Kg Shot Put (Standing), 500gms Javelin Throw (10m Approach), Pentathlon (60m · 80mH · LJ · SP · 600m).

3. **NIDJAM Selection Quota (Per-District Contingent Limits)**
   - **Total District Contingent**: Maximum **13 Athletes** and **2 Officials** selected by MDAA for NIDJAM.
   - **U-14 Category Limit**: Maximum **4 Athletes** per district (Boys and Girls together).
   - **U-16 Event Limit**: Maximum **2 Athletes** per event from the district.

4. **Entry Window & Deadlines**
   - **Online Entries Open**: 15 October 2026
   - **Online Entries Close**: 30 October 2026 (23:59:59 IST)
   - No entries accepted via WhatsApp or Email.

5. **Parent / Guardian Undertaking**
   - Mandatory consent and undertaking regarding safety and liability declaration as stipulated in Circular No. 88.

6. **District Eligibility Documents**
   - District Meet Results, Press Cuttings, Meet Photographs, District Association PAN Card.

---

### Running Locally

```bash
# Install dependencies
npm install

# Start local server
npm start
# or with file watching:
npm run dev
```

Visit [http://localhost:3000](http://localhost:3000) in your browser.

- **Athlete View**: Form validation enforces age limits, AFI UID, quota tracking, single-event constraints, and parent undertakings.
- **Admin View**: Sign in with `admin@mysuru-athletics.com` and passcode `mysuru2026`. View compliance checklist, entries table, and export AFI CSV.

---

### Deploying to Vercel

The application is structured for instant zero-configuration deployment to Vercel:
- **Frontend**: `index.html` and static assets served on Vercel's Edge CDN.
- **API**: Serverless functions in `api/` (`api/config.js`, `api/stats.js`, `api/entries.js`).
- **Routing**: `vercel.json` rewrites `/config.js` to `/api/config`.

#### Method 1: Deploy using Vercel CLI (Fastest)

```bash
# In the project directory, run:
npx vercel

# Follow the interactive prompts:
# - Set up and deploy? Yes
# - Which scope? (Select your account)
# - Link to existing project? No
# - Project name? mysuru-athletics-entry
# - Directory? ./
# - Want to modify settings? No

# Deploy to production:
npx vercel --prod
```

#### Method 2: Deploy using GitHub & Vercel Dashboard

1. Push your repository to GitHub.
2. Go to [https://vercel.com/new](https://vercel.com/new) and import the repository.
3. In **Environment Variables**, add the variables from `.env`:
   - `SUPABASE_URL`
   - `SUPABASE_ANON_KEY`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - `ADMIN_EMAIL`
   - `ENTRY_DEADLINE`
   - `ENTRY_OPEN_DATE`
4. Click **Deploy**.
