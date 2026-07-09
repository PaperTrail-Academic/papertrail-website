# Draft FAQ — targets the searches above, in brand voice

For review only — not yet added to the site. Suggested home: a new FAQ section on the homepage (best for SEO weight) or its own `/faq/` page. Written to match your existing register: teacher-to-teacher, honest about limits, no verdicts.

---

**Is PaperTrail an AI detector?**

No. AI detectors return a probability — a percentage chance that text was AI-generated — and that number is often wrong, especially for ESL students and neurodivergent writers whose natural phrasing can read as "unnatural" to a model. PaperTrail doesn't guess at authorship from the text alone. Inspect shows you the document's actual revision history. StyleMatch compares writing style against a student's own past work. Verify adds AI-assisted analysis on top of that evidence — but the output is always a report you interpret, never a score that decides for you.

**Is PaperTrail a plagiarism checker?**

No — that's a different problem. Plagiarism checkers like Turnitin scan submitted text against a database of existing sources to find copied passages. PaperTrail doesn't do that. Verify Citations checks something more specific: whether the sources a student cites actually say what the essay claims they say. That catches fabricated or misrepresented evidence, which a plagiarism checker won't.

**How can I tell if a student used ChatGPT to write an essay?**

No single signal proves it either way, which is why we built three independent ones instead of one verdict. Inspect shows you the session history — how the document was actually typed, including any large paste events. StyleMatch flags whether the submitted essay's sentence structure, vocabulary, and punctuation patterns diverge from that student's own controlled writing sample. Verify adds a deeper AI-assisted read of internal consistency, and Oral Defense lets the essay generate its own follow-up questions for the student to answer out loud. None of these alone is proof. Together, they're a documentable, defensible basis for a conversation.

**Why do AI detectors give false positives?**

Because they're built to output a probability, not a fact — and that probability is trained on patterns that don't hold evenly across writers. Formulaic, tightly-structured writing (the kind ESL students and many neurodivergent students are taught to produce) frequently scores as "more likely AI" even when it's entirely their own. That's a structural problem with score-based detection, not a tuning issue. It's the reason PaperTrail doesn't produce a single AI-probability number — we show you what happened in the document instead of asking you to trust a model's guess about the text.

**How is this different from Draftback or Revision History?**

Inspect does the same core thing — replay a Google Doc's edit history — and it's free, like Process Feedback, rather than a paid subscription like Draftback. Where Inspect goes further is in what it surfaces automatically: paste-and-novelty scoring, struggle-moment detection, and a printable Process View report you can hand a student or parent, not just a scrubber timeline you watch by hand.

**What's the difference between StyleMatch and an AI detector?**

An AI detector asks "does this text look machine-generated?" StyleMatch asks a narrower, more answerable question: "does this sound like the same person who wrote their in-class sample?" It's authorship consistency, not AI-generation probability — eight stylometric metrics (function words, punctuation patterns, sentence rhythm, and more) computed client-side, nothing transmitted. That distinction matters: a ghostwritten-by-a-friend essay and a heavily-AI-edited essay can both fail a style match without either one tripping an AI detector.

**Can PaperTrail help a student prove they didn't cheat?**

Yes — that's half the point. A revision history that shows steady, incremental work over several sessions is real evidence, and it goes both directions. Inspect's Process View report and Oral Defense's recorded answers give a student something concrete to point to, not just a denial.
