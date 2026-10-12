# Mysuru District Athletics Association — Athlete Entry Portal
## 20th National Inter District Junior Athletics Meet 2026 (NIDJAM)
**Strict Compliance with AFI Circular No. 88/09-2026 (09 September 2026)**

### Overview
This portal manages athlete entries from Mysuru District for the **20th National Inter District Junior Athletics Meet 2026 (NIDJAM)** to be held at **Woxsen University, Kamkole, Sadasivpet, Hyderabad, Telangana** from **14 – 16 November 2026**, organized by the Athletics Federation of India (AFI) in collaboration with the Government of Telangana.

---

### Strict Enforced Rules (Circular No. 88/09-2026)

1. **AFI UID (Mandatory)**
   - Every athlete must register on the AFI Website via [https://ors.indianathletics.in/#/login](https://ors.indianathletics.in/#/login) and provide a valid AFI UID.
   - Digital certificates are issued strictly in accordance with the athlete's name registered in online entry.

2. **Age Brackets & Cut-off Dates (as on 16 November 2026)**
   - **Under 14 (U-14)**: Born between **17/11/2012** & **16/11/2014** (Minimum age 12 years and above as on 16 Nov 2026; boys and girls under 12 are strictly ineligible).
   - **Under 16 (U-16)**: Born between **17/11/2010** & **16/11/2012**.
   - Original Birth Certificate (issued by Municipal Corporation / Nagar Palika / District Office) or 10th Standard Certificate mandatory for Age Verification on **13 November 2026**.
   - Over-aged athletes forfeit T.A., food, and accommodation.

3. **Event Rules & Restrictions**
   - **Under 14**: Must choose exactly **one Triathlon group** (A, B, or C) + **Kids Javelin Throw** (with 5m runway, mandatory for all U-14 athletes):
     - *Triathlon A*: 60m, Long Jump (5m approach), High Jump (Scissor only)
     - *Triathlon B*: 60m, Long Jump (5m approach), Back Throw (1kg Shot Put)
     - *Triathlon C*: 60m, Long Jump (5m approach), 600m
     - *Kids Javelin Throw*: 5-meter runway (Auto-included)
   - **Under 16**: Can enter **EXACTLY ONE EVENT**:
     - *Boys*: 60m, 80m Hurdles (0.838m), 600m, High Jump (Scissor Only), Long Jump (5m Approach), 4Kg Shot Put (Standing), 600gms Javelin Throw (10m Approach), Pentathlon (60m · 80mH · LJ · SP · 600m).
     - *Girls*: 60m, 80m Hurdles (0.762m), 600m, High Jump (Scissor Only), Long Jump (5m Approach), 3Kg Shot Put (Standing), 500gms Javelin Throw (10m Approach), Pentathlon (60m · 80mH · LJ · SP · 600m).

4. **District Quotas (Per District Limits)**
   - **Total District Contingent**: Maximum **13 Athletes** and **2 Officials** (AFI free boarding and lodging cap).
   - **U-14 Category Limit**: Maximum **4 Athletes** per district (Boys and Girls together).
   - **U-16 Event Limit**: Maximum **2 Athletes** per event from the district.

5. **Entry Window & Deadlines**
   - **Online Entries Open**: 15 October 2026
   - **Online Entries Close**: 30 October 2026 (23:59:59 IST)
   - Strict no late entry rule; no entries accepted via WhatsApp or Email.

6. **Parent / Guardian Undertaking**
   - Mandatory consent and undertaking regarding mishap/incident liability declaration as stipulated in Circular No. 88.

7. **District Eligibility Documents**
   - District Meet Results, Press Cuttings, Meet Photographs, District Association PAN Card, Association Bank Account for Rail Fare reimbursement.

---

### Running the Application

```bash
# Install dependencies
npm install

# Start local server
npm start
# or for live reloading:
npm run dev
```

Visit [http://localhost:3000](http://localhost:3000) in your browser.

- **Athlete View**: Form validation enforces age limits, AFI UID, quota tracking, single-event constraints, and parent undertakings.
- **Admin View**: Sign in with `admin@mysuru-athletics.com` and passcode `mysuru2026`. View compliance checklist, entries table, and export AFI CSV.
