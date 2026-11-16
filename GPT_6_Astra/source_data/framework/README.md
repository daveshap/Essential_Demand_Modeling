# PaidToBeHuman

## Canonical production methods

Use [METHODS.md](METHODS.md) and `method-sources.json` before an authorized script-writing, PVC or animation stage. The selected utility repo is the single source of truth for that function; verify any installed copy against current canonical contents. Methods copy one way into this project. Only reusable animation assets may be exported back. Keep function instructions isolated, including title guidance versus script prose.

David Shapiro's new course and book on essential demand: why people will still pay for a human when machines can do the work, and how ordinary people can build a living on it. The book's working title is *Paid to Be Human*. The course's own title isn't decided yet, so the scripts say "this course".

The course has 17 lessons of about 15 to 20 minutes each, read by David or his professional voice clone. `course/Course plan.md` has the lessons and the format. New video-script writing and substantive revisions now live in `C:/ScriptWriter_Agent`; this repo retains approved scripts and production work.

## The lessons

Each lesson is a Word script, `course/Lesson NN - <title>.docx`, laid out like his voice-clone recording scripts. Its earlier version was built from the script as written, `course/as-written/lesson-NN-<slug>.md`, and its research file, `course/research/Lesson NN - research.md`, records the drafting research. Its contents describe that draft; they do not impose a source or review requirement.

As of 7 Nov 2026 all 17 are first drafts, written by Claude from the brief and the reference material here, for David to edit. One Claude session wrote each lesson, and the lead session then read all 17 and edited them by hand: titles and subtitles, varied openings and sign-offs, lines repeated from one lesson to the next, names reused across lessons, and facts against each lesson's research file.

On 8 Nov 2026 David went over all 17 and edited the Word files by hand. The Word files are now the scripts; `as-written/` holds the drafts from before his edits.

Each lesson's audio is read by his voice clone from its Word file: `voice/README.md` has how.

## What's where

| Folder | What |
|---|---|
| `course/` | `Course plan.md`; the 17 Word scripts; `as-written/` (each script as Markdown, the draft the Word file was built from); `research/` (historical drafting research); `lessons.json` (each lesson's script and the note printed under its title, for the builder); `ElevenLabs/` (each lesson's text as the voice clone reads it, and in `Audio/` its takes and MP3) |
| `animation/` | `PLAYBOOK.md` for lesson animation and first-job tool setup; `ASSETS.md` for the production asset catalog. The engine and renderer are not installed yet. |
| `style/` | The 12 supplied visual-reference images for Paid to Be Human |
| `voice/` | Project adapter `canonical-pvc.toml` and handoff to `C:/ProfessionalVoiceClone`; original reader/settings/packages retained as historical production snapshots; `README.md` has the current commands |
| `punchroll/` | The measuring and mastering package the reader uses: the speech recognizer, the master and the ACX check |
| `source/` | The brief: `Paid_to_Be_Human_Conversation_Export.md`, the conversation in which David developed the course, ending with the 17-lesson outline |
| `reference/background/` | His white paper, *Post-Labor Economics: Derived vs Essential Demand* v01, as the Word file and as a text extraction with its tables inline, for searching and for line references |
| `reference/style-guide/` | Migration pointer to ScriptWriter_Agent, plus historical voice/style evidence; no active authoring workflow |
| `reference/scripts-playbook/` | Handoff to ScriptWriter_Agent and retained subject definitions for existing research links |
| `reference/writing/` | Historical writing research and example scripts; no checker comparison or production authority |
| `reference/transcripts/` | His finished recordings, transcribed: how he actually talks |
| `reference/youtube-scripts/` | His YouTube scripts as written, the Word files he read from, and their research, among them "Essential Demand" |
| `reference/labor-zero-course/` | His previous course, the Post-Labor Economics mini course: its plan, Word scripts, as-written and research files, the final texts the clone read, the changes between them, and the clone's formatting guide |
| `tools/` | Retired authoring entry points that identify ScriptWriter_Agent; the original Word template remains as an archive |
| `docs/` | `provenance.csv`: historical import locations, sizes and hashes |

## Paths in the copied documents

The copied historical documents retain older paths. This table locates retained production/reference copies; it does not reactivate their writing workflows. Active video authoring is in ScriptWriter_Agent:

| The documents say | Here |
|---|---|
| `scripts/style-guide/` | `reference/style-guide/` |
| `scripts/PLAYBOOK.md` | `reference/scripts-playbook/PLAYBOOK.md` |
| `scripts/README.md`, `scripts/as-written/`, `scripts/research/`, `scripts/recording-scripts/` | `reference/youtube-scripts/` and its folders of the same names |
| `scripts/recording-scripts/Recording Script - <Title>.docx`, a finished script's Word file | for this course, `course/Lesson NN - <title>.docx` |
| `recordings/transcripts/` | `reference/transcripts/` |
| `studio/writing/research.md`, `studio/writing/models/` | `reference/writing/research.md`, `reference/writing/models/` |
| `studio/writing/tools/` or `writing/tools/` (`check_script.py`, `build_oneoff.py`) | `tools/` |
| `writing/fire-notes.md`, in the Labor Zero plan | `reference/style-guide/Fire notes.md` |
| `course/ElevenLabs/`, in the Labor Zero formatting guide | `reference/labor-zero-course/elevenlabs/` |

Not copied, because this course doesn't need them: `recordings/outtakes/`, `recordings/tools/`, the Labor Zero formatting check (`scripts/course_format_check.py`, which asks for that course's audio tags) and its `course/tests/`. That course's reader, `scripts/course_lesson.py`, was copied on 8 Nov 2026 into `voice/scripts/`.

## How to work here

For animation, use [animation/PLAYBOOK.md](animation/PLAYBOOK.md). It starts from the existing lesson MP3 and current edited script, follows the supplied style references, and applies the shared asset-library and work-limit rules. The first animation job includes setting up only the local tools it needs; this documentation update does not install a renderer or produce a film.

Python is `py -3.13`, with `python-docx`.

For script writing, start a session in `C:/ScriptWriter_Agent`. Its portable tools format new manuscripts there; the local authoring commands now exit with that destination. The imported project map identifies the edited Word baseline and historical Markdown. Production files here are not automatically overwritten by a writer build.

For authorized new narration, use the canonical shared tool with local job paths (`voice/README.md`). These are stage examples, not instructions to regenerate completed lessons:

```
py -3.13 C:/ProfessionalVoiceClone/pvc.py prepare "<assigned lesson>" --config voice/canonical-pvc.toml
py -3.13 C:/ProfessionalVoiceClone/pvc.py lesson make "<assigned lesson>" --config voice/canonical-pvc.toml --max-chars 30000
```

## Where it came from

The source brief and reference copies came from the imports recorded in `docs/provenance.csv`. Those rows record historical file states; they do not require a ledger entry or new hash for every edit. The voice reader and measuring helpers were copied from Labor Zero and adapted here. This project's current instructions govern their use. The source projects were not changed by those copies. `voice/.env` contains the ElevenLabs credentials and is excluded from Git.
