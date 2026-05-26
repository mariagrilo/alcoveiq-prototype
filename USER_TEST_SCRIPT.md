# AlcoveIQ — User Test Script

**Duration**: ~50–60 min
**Context**: Stakeholder + real end user, remote, screen-sharing

---

## 0 · Pre-session (5 min before)

- [ ] Open `localhost:8765/` in Chrome
- [ ] Resize Chrome to typical laptop size (~1280×800)
- [ ] Reset to **populated demo** (avatar → Reset current demo) — clean slate
- [ ] Start screen recording
- [ ] Open a notes doc with the sections below

---

## 1 · Intro (≤ 3 min)

> "This is AlcoveIQ — a portfolio entity management platform we're prototyping. I want to learn how you'd actually use it. **Please think out loud** — what you notice, what confuses you, what you'd click. Some things might not work; that's fine, just tell me what you expected.
>
> There are no right answers. I'm watching the product, not you."

**Don't lead. Don't explain features. Let them discover.**

---

## 2 · Empty workspace — first-time user (15 min)

**Setup**: avatar → Load empty demo → confirm.

**Prompt**:
> "Imagine you just signed up. Tell me what you see and what you'd try first."

### Observe (don't interrupt)
- [ ] Where do their eyes go first?
- [ ] Did they notice the chat? Try it?
- [ ] Did they go to Entities? Did the empty-state CTA make sense?
- [ ] Did they try the formation wizard, the chat, or both?
- [ ] If they used the **chat to form an entity** — did the multi-turn flow feel intelligent or robotic?
- [ ] If they typed a name with the wrong suffix — how did they react to the warning?
- [ ] Did they notice the new entity's future compliance items appear on Compliance?

### Ask (after they explore for ~7 min)
- "What was your first impression?"
- "If this were a real product, what would you do next?"
- "Was there anything you expected to find but couldn't?"

---

## 3 · Populated workspace — existing user (20 min)

**Setup**: avatar → Load populated demo → confirm.

**Prompt**:
> "Now imagine you've been using AlcoveIQ for a year. This is your portfolio. Walk me through how you'd start a typical Tuesday."

### Observe
- [ ] Did they land on Today and scan top-to-bottom?
- [ ] Did **Blockers / Overdue / Due this week** sections make sense?
- [ ] Did they click File now anywhere? Did the order modal flow feel natural?
- [ ] Did they hit the **Reinstate RA blocker chain** (Maple Holdco)? Did they understand the dependency?
- [ ] Did they see the **"Needs your input"** section?
- [ ] Did they trust the **payment confirmation modal**? Did they notice the amount in the button?
- [ ] Did they try the **conversational AI** to file something / look something up?
- [ ] Did they notice orders progress automatically (In Progress → Sent to RA → Completed)?
- [ ] Did they see new documents appear after completion?

### Ask
- "On a busy Tuesday, what would you check first here?"
- "What's the difference between Today and Compliance, in your mind?"
- "When would you talk to the AI vs click around?"

---

## 4 · Targeted UX probes (10 min)

Only ask these if they didn't surface naturally during free exploration.

### Naming & IA
- "What does **'Today'** mean to you here? Would you call this something else?"
- "Where would you go to see your **org chart**? When would you actually need it — exporting it, browsing it, both?"
- "Does **'+ New order'** in the top right make sense? When would you click it?"

### AI assistant
- "Try asking the AI something weird or unexpected. What do you wish it could do?"
- "Would you prefer the AI on the side (where it is) or somewhere else?"
- "When you ask the AI to form an entity, does it feel like it's helping or interrogating you?"

### Trust & confidence
- "If the AI filed something automatically, would you trust it? What would you need to see to trust it?"
- "Does the **payment confirmation** make you feel safe approving a charge?"
- "When something says 'Auto-filed', does that worry you?"

### Empty / blocked / error states
- "Look at the Maple Holdco rows that say 'Locked' — does it tell you enough about why?"
- "What happens if an order fails? Where would you expect to see that?"

---

## 5 · Wrap-up (8 min)

- "Top 3 things you liked."
- "Top 3 things you'd change."
- "Would you use this? In what role on your team?"
- "What's it missing compared to your current tool (Harbor / NetSuite / Excel)?"
- "On a scale of 1–10, how much would you trust this with real compliance data today?"
- "What's the one thing that would make this a no-brainer to adopt?"

---

## Cheat sheet — UX bets to validate

Use this AS A REFERENCE if a topic doesn't come up organically. Don't read it to them.

| Bet | Validation question |
|---|---|
| **Action-first Today** is better than a status dashboard | Did they look at Today first? Did they act from it? |
| **Conversational AI** is realistic | Did they trust it? Did they hit its limits gracefully? |
| **Multi-turn entity creation** via chat is useful | Did they prefer chat or wizard? Why? |
| **Blocker chain UX** is comprehensible | Did they understand "Reinstate RA first" before File now? |
| **Auto-filed / automatic progression** is acceptable | Did they trust it? Would they want manual confirmation? |
| **"Needs your input"** state earns its place on Today | Did they value it or want it elsewhere? |
| **Payment confirmation** earns user trust | Did they read the amount/recipient before clicking Pay? |
| **Today vs Dashboard naming** is honest | Did "Today" cause confusion? |
| **Org Chart placement** matches use case | Browse view, export-only, or both? |
| **AI chat at 1/3 width** is right | Did they collapse it? Resize it? |

---

## Post-session

- [ ] Stop recording
- [ ] Tag the top 3 surprises immediately while fresh
- [ ] Note any "I don't know how to..." moments — those are gold
- [ ] If they suggested a feature, note whether it was *missing* or *misnamed*
- [ ] Quick triage: must-fix / should-fix / nice-to-have before sharing with team
