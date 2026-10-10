#!/usr/bin/env python3
r"""commit_trailer_check.py -- fail a commit range that carries an AI co-author trailer.

Created: 2026-10-06
Updated: 2026-10-10 (prompt 13 of the 2026-10-07 queue: a commit's trailers are git's own, read with git log -z and %(trailers:unfold,only) in the call that reads the range, and a message that is not yet a commit is read by git interpret-trailers --parse --unfold --no-divider, so the nine shapes git reads as a model trailer that this file passed are findings, a 0x1E byte no longer truncates a record, a key is judged with the whitespace before its colon removed, and a value opening with a non-breaking space is still judged; a forced push to the default branch whose old tip no clone holds exits 2 naming the push, and a new branch reads its head minus origin/<--default-branch> only); 2026-10-06 (the commit-msg hook's gpt- signal and plural co-author keys join AI_VENDORS and ATTRIBUTION_KEYS, so the hook can judge with this pattern without losing either; earlier the same day a "(cherry picked from commit" line continues the final trailer block, as git's own trailer parser counts it, so a model trailer that git cherry-pick -x carries over is read; earlier the same day prompt 10 of the 2026-09-28 tools queue: new; the trailer pattern moved here verbatim from cowork_util.py's audit-commit-trailer section, so that gate and every repo's CI judge a message with one copy of it)

Usage, as every repo's .github/workflows/commit-trailer.yml runs it:

    python3 .github/scripts/commit_trailer_check.py --repo . --base <before> --head <after> --ref <ref> --forced <true|false> --default-branch <branch>
    python3 .github/scripts/commit_trailer_check.py --repo . --base <base sha> --head <head sha>

Exit 0 = no commit in the range carries one, with the commit count printed
first. Exit 1 = at least one does, each named by SHA, author and line. Exit 2 =
the range could not be read, which is never a pass. Standard library only, so
the same bytes run on every runner without an install step.

THE INCIDENT. lazygrip 85032aa, "Add normalizeCollectionEntries helper for
malformed collection_sequences values", was authored 2026-09-30 14:20:41 -0500
with a model Co-Authored-By line. Its committer is GitHub with a web-flow
signature and no pull request: it was made on GitHub itself, so no commit-msg
hook in any clone saw its message. Its two check runs, schema-from-scratch and
build, passed, because lazygrip's ci.yml reads no commit message, and it
surfaced on 2026-10-04 only because a Cowork session ran audit-pre-handoff. The
one CI step that read a commit message was GRIP-Tools' own; the other twelve
repos that push to GitHub read none. This file is what their CI now runs.

ONE FILE, THREE READERS. The source lives at tools/hooks/commit_trailer_check.py
in GRIP-Tools. cowork_util.py loads it from there and audit-commit-trailer judges
every commit through its functions, hooks/commit-msg judges with its lists held
equal by tests/audit_subcommands/test_commit_msg_hook.py, so the gate that reads the whole workspace
after the fact and the CI step that reads one pushed range cannot disagree about
what a trailer is. Every other repo carries a byte copy at
.github/scripts/commit_trailer_check.py, and audit-commit-trailer-ci fails when a
copy differs from this file. Edit this file, then re-copy it; never edit a copy.

WHY tools/hooks/. It is the same guard as hooks/commit-msg at a later moment:
the hook refuses the trailer at commit time in a clone, and this refuses it at
push time on GitHub, where a web-made commit never met a hook. audit-commit-hooks
reads its template from this directory and audit-commit-trailer-ci reads both of
its sources from here, the workflow template hooks/commit-trailer.yml included.
cowork_util.py loads it by path from the directory holding cowork_util.py, never
through GRIP_TOOLS_ROOT, so a test sandbox that moves the tools root cannot move
the pattern out from under the gate.

THE RANGES, each a measured fact rather than a guess:
  * a push hands its before and after SHAs, and the range is before..after;
  * a push that CREATES a branch carries an all-zero before SHA, and the range
    is every commit reachable from the pushed SHA and not from the remote's
    default branch, origin/<--default-branch>. Until 2026-10-10 it was "and
    from no other branch of the remote", and two branches created by one push
    (git push origin x:refs/heads/a x:refs/heads/b) each saw the other as an
    existing branch holding x, so both ranges were empty and both printed
    "PASS (nothing asserted)". A push that creates the default branch itself,
    a repository's first push, reads the whole branch;
  * a FORCED push whose before SHA no clone holds -- a rebase pushed with
    --force-with-lease, as Dependabot does to its own branches -- takes the
    new-branch rule, because GitHub marks the push forced (--forced true); the
    2026-10-06 review measured every such push failing at exit 2 without it.
    A forced push of the DEFAULT branch whose old tip no clone holds exits 2
    naming the push instead: the new-branch rule would read all of that
    branch's history, which on lazygrip is its four known hits and on a fresh
    ems or hub runner the history before 2026-09-16;
  * a pull request hands its base and head SHAs, and the range is base..head.
A range needs the history under it, so the workflow checks out with
fetch-depth 0; actions/checkout@v7 defaults to 1, measured in the runner's cached
copy of the action, and a depth-1 clone cannot resolve a before SHA at all.

THE TRAILERS ARE GIT'S OWN since 2026-10-10. Until then this file read a
message's final block with its own line rule, and measured that day against
git 2.55.0.windows.3 it passed nine shapes git reads as a model trailer: a
comment line, an indented continuation or the vendor word on a continuation
below the model line, a space before the colon, an empty "Fixes:" below it, a
non-breaking space after the colon, a Signed-off-by followed by a prose line
(git's 25 percent rule), a scissors line kept by -F, and a 0x1E byte anywhere in
the message, which truncated the record before the body was read. A commit's
trailers now come from git log's %(trailers:unfold,only) in the call that reads
the range, and a message that is not a commit yet from git interpret-trailers,
so the commit-msg hook, this checker and audit-commit-trailer read a trailer
the way git does and cannot disagree with it or with each other.
"""

import argparse
import os
import subprocess
import sys

# ------------------------------------------------------------------
# THE PATTERN. Moved verbatim from cowork_util.py on 2026-10-06, comments and
# all, because the comments carry the measurements that justify each literal.
# "This gate" in them is audit-commit-trailer, which reads its pattern from this
# file now; the 2026-09-16, 2026-09-18 and 2026-09-19 decisions stay recorded
# in the block comment above that gate's banner in cowork_util.py.
# ------------------------------------------------------------------

# MEASURED. The literal substrings that actually occur in the 82-commit
# population, counted 2026-09-16: four distinct values totalling 82 occurrences --
# "Claude Opus 5" at 66, "Claude Opus 4.8" at 12, "Claude Opus 4.8 (1M context)"
# at 3 and "Claude Opus 4.7" at 1 -- every one of them carrying the address
# noreply@anthropic.com. Two substrings cover all four values and the address.
AI_MEASURED = ("anthropic", "claude")

# FORWARD-LOOKING. None of these has EVER appeared in this workspace; the
# 2026-09-16 sweep found zero occurrences of any of them. They are here so a
# different vendor's trailer is a finding on the day it first lands rather than
# after its own 82-commit population has accumulated. They are LITERALS rather
# than a heuristic, which is the whole reason this subject is gateable at all:
# an intent-matching rule over commit prose would be arguing with English, while
# a substring over a trailer VALUE is a fact about bytes.
#
# "gpt-" JOINED ON 2026-10-06, from the commit-msg hook. hooks/commit-msg had
# refused a `\bgpt-` value since 2026-09-18 while this list let one through, and
# aligning the hook with this pattern must not drop a literal it already
# enforced. A GPT model name with no vendor word, "GPT-5 <bot@...>", was the shape
# only the hook caught. Scored that day over the full history of the 14 owned
# repos, 3487 commits: it adds no finding.
AI_VENDORS = ("openai", "chatgpt", "copilot", "gemini", "codex", "devin", "aider", "gpt-")

# WHICH LINES ARE TRAILERS IS GIT'S ANSWER, NOT THIS FILE'S, since 2026-10-10.
# Until then a KEY_RE here, ^[A-Za-z][A-Za-z0-9-]*:[ \t]*\S, decided which lines
# extended the final block, and three of the nine shapes git reads as a model
# trailer failed it outright: "Co-Authored-By : ..." has a space before the
# colon, a bare "Fixes:" below the model line has no value, and a non-breaking
# space after the colon is whitespace to Python's \S, so the block ended there.
# See read_range and message_trailers for where the trailers come from now.

# THE FINDING IS AN ATTRIBUTION TRAILER, AND THIS NARROWING IS MEASURED RATHER
# THAN GUESSED. The final-block parse alone is not sufficient, and the live tree
# proved it on the first real run: a commit message whose whole body is one
# paragraph beginning `Description: ...` is trailer-shaped by the block rule and
# by git's own parse -- `git interpret-trailers --parse` returns that line as a
# trailer -- so a signal anywhere in that prose reads as an attribution.
#
# Measured over the FULL history of every owned repo under C:\Dev, 2026-09-16:
# 86 final-block lines carry a signal literal. 82 of them are Co-Authored-By --
# exactly the population this gate was specified against, reproduced to the
# commit. The other FOUR are prose paragraphs under the keys Description,
# Verifications, Updated and HOUSEKEEPING: claudefix 1cb6f0b naming
# `Claude/Validate-Scripts.ps1` and `ClaudeFix.zip`, and three EMS bodies. Every
# one of them is the same family as the two false positives the final-block rule
# was written to kill -- a filename or folder reference, not an attribution --
# and one of them lives in a repository whose own NAME contains the signal, so
# it would recur for as long as that repo is committed to.
#
# THE DISCRIMINATOR IS THE KEY, AND IT IS STRUCTURAL RATHER THAN A WORD LIST.
# Every attribution trailer git or any tool has ever written ends in `-by`:
# Co-Authored-By, Signed-off-by, Reviewed-by, Tested-by, Acked-by, Helped-by,
# Reported-by, Suggested-by, and any Assisted-By or Generated-By a future vendor
# invents. That suffix is the rule; `author`, `co-author` and `cc` are named
# outright because they are the three attribution keys that do not carry it.
# Applying it to the 86 leaves exactly the 82 and drops exactly the 4.
#
# `-with` IS PART OF THAT SUFFIX RULE, AND THE HOLE WAS LIVE-FIRED RATHER THAN
# IMAGINED (2026-09-16). The sentence above ends "any Assisted-By or
# Generated-By a future vendor invents" -- and the shape that slips through is
# exactly one character class away from it. Scored against the SHIPPED gate, a
# throwaway repo outside C:\Dev carrying one commit whose entire final trailer
# block is `Generated-with: <model> <noreply@anthropic.com>` EXITS 0: the key
# does not end in `-by` and is none of the four named keys, so the line is
# never judged at all and a trailer naming a model passes silently.
#
# CLOSING IT IS FREE, AND THAT IS MEASURED RATHER THAN ARGUED. Re-scored over
# the same population as above -- the full history of all fifteen owned repos
# under C:\Dev, 3048 commits, the 86 final-block lines carrying a signal
# literal:
#     shipped rule,  `-by`            + the four named keys -> keeps 82
#     widened rule, (`-by`|`-with`)   + the four named keys -> keeps 82
# Identical. Not one of the four measured prose false positives -- the
# Description, Verifications, Updated and HOUSEKEEPING lines -- is dragged back
# in, because none of THOSE keys ends in `-with` either.
#
# THE SUFFIX IS STILL THE RIGHT INSTRUMENT, and widening it is why. A longer
# word list would have to name Generated-with, Created-with, Built-with,
# Written-with and whatever the next vendor ships, one commit behind each; the
# suffix pair covers the whole `<verb>-<preposition>: <agent>` family at zero
# measured cost, which is precisely the property that made `-by` the rule.
ATTRIBUTION_KEY_SUFFIXES = ("-by", "-with")
#
# WHAT THIS DELIBERATELY DOES NOT DO is ban the bare trailer. Co-Authored-By
# naming a human being is a legitimate co-author and passes: the key decides
# whether the line is an attribution at all, and the VALUE decides whether the
# attribution is to a machine.
#
# THE PLURAL KEYS JOINED ON 2026-10-06 for the reason "gpt-" did: the commit-msg
# hook's `co-?authors?` had refused Co-Authors and Coauthors since 2026-09-18,
# and alignment must not drop them. Scored over the same 3487 commits, they add
# no finding.
ATTRIBUTION_KEYS = ("author", "co-author", "coauthor", "cc", "co-authors", "coauthors")

# A LINE GIT WRITES INTO THE TRAILER BLOCK WITH NO KEY, measured 2026-10-06 on
# git 2.55.0.windows.3: `git cherry-pick -x` appends "(cherry picked from commit
# <sha>)" straight under a message's trailer block, and git's own trailer parser
# counts it as part of the block, so a Co-Authored-By above it is still a
# trailer. This file carried its own continuation rule for that line until
# 2026-10-10; git's parser carries it now, and tests t14 and t15 still hold it.


def trailer_key(line):
    """The KEY of a trailer line, lowercased, with the whitespace before its colon removed.

    git prints every trailer it parses as "Key: value", the key already
    trimmed, measured 2026-10-10 on git 2.55.0.windows.3 for "Co-Authored-By :"
    and "Co-Authored-By<TAB>:" alike. The key is still trimmed here, so the
    judgement of a trailer line does not lean on that formatting: a caller
    handing a raw "Co-Authored-By : <model>" line gets the same answer git's
    normalised one gets.
    """
    return line.partition(":")[0].strip().lower()


def is_attribution(line):
    """True when `line`'s trailer KEY attributes authorship to somebody."""
    key = trailer_key(line)
    return key.endswith(ATTRIBUTION_KEY_SUFFIXES) or key in ATTRIBUTION_KEYS


def trailer_value(line):
    """The part of a trailer line after its FIRST colon, lowercased and stripped.

    The KEY is deliberately not tested. `Co-Authored-By` carries no signal in
    either direction, and testing it would make the gate fire on a human
    co-author whose surname happened to collide with a vendor literal.
    """
    _key, _sep, value = line.partition(":")
    return value.strip().lower()


def signals(value_lower):
    """Every signal literal present in `value_lower`, measured set first."""
    return [s for s in (*AI_MEASURED, *AI_VENDORS) if s in value_lower]


def trailer_findings(trailers):
    """[(line, value, hits)] for each trailer in `trailers` that attributes authorship to a machine.

    THE WHOLE JUDGEMENT OF ONE MESSAGE, in one function every reader calls.
    `trailers` is git's own list for that message, one "Key: value" line per
    trailer, unfolded: parse_log hands it over for a commit and message_trailers
    for a message that is not a commit yet. A trailer is a finding when its key
    is an attribution and its value carries a signal. audit-commit-trailer
    applies its declared exceptions to what this returns, and the CI check fails
    on anything it returns.

    A VALUE OPENING WITH A NON-BREAKING SPACE IS JUDGED like any other. git keeps
    the U+00A0 in the value it returns, measured 2026-10-10, and the value is
    read as a substring, so nothing about the character after the colon decides
    whether the line is a trailer; that is git's answer, already given.
    """
    found = []
    for line in trailers:
        if ":" not in line:
            continue
        if not is_attribution(line):
            continue
        value = trailer_value(line)
        hits = signals(value)
        if hits:
            found.append((line, value, hits))
    return found


# ------------------------------------------------------------------
# GIT. Everything below runs git and nothing above does.
# ------------------------------------------------------------------

# ONE git log CALL READS THE RANGE AND ITS TRAILERS, keyed by commit: the SHA,
# the subject, then git's own trailer list, unfolded, one per line. -z ends each
# record with a NUL. Until 2026-10-10 the records ended in 0x1E through
# %x1e, and a 0x1E byte anywhere in a message ended its record early, so a model
# trailer below it was never read (measured that day: ai_attribution_lines over
# the body found it, the CLI over the range did not). git commit refuses a NUL
# in a message ("a NUL byte in commit log message not allowed", measured
# 2026-10-10), so a NUL cannot end a record early.
LOG_ARGS = ("-z", "--format=%H%n%s%n%(trailers:unfold,only)")
RECORD_SEPARATOR = "\x00"

# THE PARSER FOR A MESSAGE THAT IS NOT A COMMIT YET, the commit-msg hook's and
# ai_attribution_lines', run over the message with no repository needed.
# MEASURED 2026-10-10 on git 2.55.0.windows.3, a message file at a time:
#   * a comment line (core.commentChar, "#" by default) is skipped and does not
#     end the trailer block, so a model line with a comment below it is a
#     trailer, and a commented-out trailer is not one;
#   * the scissors line "# ------------------------ >8 ------------------------"
#     ends the message: nothing below it is read, a trailer below it included;
#   * a 0x1E byte is an ordinary byte, kept in the value it sits in;
#   * a "---" line ends the message too unless --no-divider is passed, and git
#     log reads a commit's trailers with no divider: a model trailer above a
#     "---" line came back from interpret-trailers without --no-divider and from
#     neither git log nor interpret-trailers with it. --no-divider is what keeps
#     the hook and the checker one parser;
#   * a message of one paragraph has no trailer, since its first paragraph is
#     the subject, and git log reads the commit made from it the same way.
# Over 24 shapes, git log's %(trailers:unfold,only) of each commit and this
# call's output for its message were the same lines.
INTERPRET_TRAILERS = ("interpret-trailers", "--parse", "--unfold", "--no-divider")


class RangeError(Exception):
    """The range could not be read. main() turns it into exit 2, never 0."""


def parse_log(payload):
    """Split git log output in the LOG_ARGS shape into (sha, subject, trailers) triples.

    `trailers` is git's list of the commit's trailer lines, empty when it has
    none. A record is cut on RECORD_SEPARATOR alone, never on a newline, so a
    subject or a trailer value holding any other byte stays whole.
    """
    commits = []
    for record in payload.split(RECORD_SEPARATOR):
        record = record.lstrip("\n")
        if not record.strip():
            continue
        sha, sep, rest = record.partition("\n")
        if not sep:
            continue
        subject, _sep, trailers = rest.partition("\n")
        commits.append((sha.strip(), subject, [ln for ln in trailers.split("\n") if ln.strip()]))
    return commits


def message_trailers(message):
    """git's trailer lines for `message`, a message that is not a commit yet. Raises RangeError.

    See INTERPRET_TRAILERS for what git does with comment lines, a scissors
    line, a 0x1E byte and a "---" line, each measured.
    """
    try:
        proc = subprocess.run(
            ["git", *INTERPRET_TRAILERS],
            input=(message or "").encode("utf-8"),
            capture_output=True,
            timeout=60,
        )
    except FileNotFoundError as exc:
        raise RangeError("git is not on PATH (%s)" % exc) from exc
    except subprocess.TimeoutExpired as exc:
        raise RangeError("git interpret-trailers timed out after 60s") from exc
    if proc.returncode != 0:
        raise RangeError(
            "git interpret-trailers exited %d: %s" % (proc.returncode, proc.stderr.decode("utf-8", "replace").strip())
        )
    text = proc.stdout.decode("utf-8", "replace").replace("\r\n", "\n")
    return [ln for ln in text.split("\n") if ln.strip()]


def ai_attribution_lines(message):
    """[(line, value, hits)] for each trailer of `message` that attributes authorship to a machine.

    For a message that is not a commit yet; a commit's trailers come from
    read_range. A message whose list is empty carries no AI co-author trailer
    by this file's definition.
    """
    return trailer_findings(message_trailers(message))


def _git(repo, *argv):
    """Run one read-only git command in `repo`. Returns (returncode, stdout, stderr)."""
    try:
        proc = subprocess.run(
            ["git", "-C", str(repo), *argv],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=300,
        )
    except FileNotFoundError as exc:
        raise RangeError("git is not on PATH (%s)" % exc) from exc
    except subprocess.TimeoutExpired as exc:
        raise RangeError("git timed out after 300s: git %s" % " ".join(argv)) from exc
    return proc.returncode, proc.stdout or "", proc.stderr or ""


def _is_zero(sha):
    """True for the all-zero SHA a push carries when it creates or deletes a branch."""
    return bool(sha) and set(sha) == {"0"}


def _resolve(repo, sha, what):
    """The full SHA of commit `sha`, or RangeError naming which end is missing."""
    rc, out, _err = _git(repo, "rev-parse", "--verify", "--quiet", sha + "^{commit}")
    if rc != 0 or not out.strip():
        raise RangeError(
            "the %s SHA %s is not a commit in this clone -- a shallow checkout, a SHA from another "
            "repository, or a force push whose old tip is on no branch any more" % (what, sha)
        )
    return out.strip()


def _branch_of(ref):
    """The branch name a pushed ref names: refs/heads/x -> x, a bare x -> x, any other refs/ -> None."""
    if ref.startswith("refs/heads/"):
        return ref[len("refs/heads/") :]
    if ref.startswith("refs/"):
        return None
    return ref or None


def range_revisions(repo, base, head, ref="", remote="origin", forced=False, default_branch=""):
    """(git log revision arguments, one-line description) for the range, or RangeError.

    base..head when base is a commit. When base is all zeros the push created
    a branch, and the range is the new-branch rule's, _branch_revisions. A
    FORCED push whose base is a commit this clone does not hold -- the old tip
    of a rebase, reachable from nothing once the push lands -- takes the same
    rule, unless it pushed the default branch; an unforced one with an unknown
    base stays unreadable.

    THE FORCED PUSH OF THE DEFAULT BRANCH IS UNREADABLE, since 2026-10-10. The
    new-branch rule reads the pushed head minus the default branch, and here the
    default branch IS the pushed head, so the rule read the branch's whole
    history: lazygrip's four known hits, or on a fresh ems or hub runner the
    history before 2026-09-16, which the workflow's own comment says is never
    read. What the push removed and what it brought cannot be told apart without
    the old tip, so the push is named and read by a person, at exit 2.
    """
    rc, _out, err = _git(repo, "rev-parse", "--git-dir")
    if rc != 0:
        raise RangeError("not a git repository: %s (%s)" % (os.path.abspath(str(repo)), err.strip()))
    if not head or _is_zero(head):
        raise RangeError("the head SHA is %r -- a range has to end at a commit" % head)
    head_full = _resolve(repo, head, "head")
    if not base:
        raise RangeError("no base SHA was given -- pass the push's before SHA or the pull request's base SHA")
    if _is_zero(base):
        return _branch_revisions(
            repo, head_full, ref, remote, default_branch, "the base SHA is all zeros, which names a new branch"
        )
    try:
        base_full = _resolve(repo, base, "base")
    except RangeError:
        if not forced:
            raise
        if default_branch and _branch_of(ref) == default_branch:
            raise RangeError(
                "a FORCED push of the default branch %s replaced %s with %s, and no clone holds the old tip, so "
                "what the push removed and what it brought cannot be told apart. Read as a new branch it would "
                "be the branch's whole history, which this check never reads. Compare the old and the new tip "
                "by hand, and re-run this job once the old tip is fetched or the push is reverted"
                % (default_branch, base, head_full)
            ) from None
        return _branch_revisions(
            repo,
            head_full,
            ref,
            remote,
            default_branch,
            "the push was forced and its old tip %s is on no branch of this clone" % base,
        )
    return ["%s..%s" % (base_full, head_full)], "%s..%s" % (base_full, head_full)


def _branch_revisions(repo, head_full, ref, remote, default_branch, why):
    """The new-branch rule: the pushed head minus the remote's default branch, origin/<default_branch>.

    THE DEFAULT BRANCH ALONE, since 2026-10-10. The rule was "minus every other
    branch of the remote", and two branches created by one push (`git push
    origin x:refs/heads/a x:refs/heads/b`, reproduced by the 2026-10-06 review)
    each saw the other as an existing branch holding x: both ranges were empty,
    both runs printed "PASS (nothing asserted)", and on hub's queue the window for
    a branch to appear in between is unbounded. Each new branch now reads its own
    commits whatever else the push created. THE PRICE IS A RE-READ: a branch
    stacked on another branch that is not merged yet reads that branch's commits
    again, which costs a second look at commits already judged and never hides
    one. A push that creates the default branch itself, a repository's first
    push, reads the whole branch, since there is nothing older to subtract.
    """
    if not ref:
        raise RangeError(
            "%s, and no --ref says which branch was pushed -- without it a push of the default branch "
            "cannot be told from a new branch" % why
        )
    if not default_branch:
        raise RangeError(
            "%s, and no --default-branch names the remote's default branch, the one a new branch is read "
            "against -- the workflow passes github.event.repository.default_branch" % why
        )
    if _branch_of(ref) == default_branch:
        whole = "%s, read as the default branch %s being created (%s): every commit reachable from %s" % (
            ref,
            default_branch,
            why,
            head_full,
        )
        return [head_full], whole
    default_ref = "refs/remotes/%s/%s" % (remote, default_branch)
    rc, out, _err = _git(repo, "rev-parse", "--verify", "--quiet", default_ref + "^{commit}")
    if rc != 0 or not out.strip():
        raise RangeError(
            "%s, and %s, the default branch a new branch is read against, is not in this clone -- check out "
            "with fetch-depth 0" % (why, default_ref)
        )
    desc = "%s, read as a new branch (%s): commits reachable from %s and not from %s" % (
        ref,
        why,
        head_full,
        default_ref,
    )
    return [head_full, "--not", default_ref], desc


def read_range(repo, base, head, ref="", remote="origin", forced=False, default_branch=""):
    """([(sha, subject, trailers)], description) for every commit in the range, or RangeError.

    THE TRAILERS COME FROM THE SAME git log CALL that lists the range, so each
    commit is judged by git's own reading of its message; see LOG_ARGS.
    """
    revisions, desc = range_revisions(repo, base, head, ref, remote, forced, default_branch)
    rc, payload, err = _git(repo, "log", *LOG_ARGS, *revisions)
    if rc != 0:
        raise RangeError("git log failed over %s: %s" % (desc, err.strip()))
    return parse_log(payload), desc


def _author(repo, sha):
    rc, out, _err = _git(repo, "log", "-1", "--format=%an <%ae>", sha)
    return out.strip() if rc == 0 and out.strip() else "(author unreadable)"


def _ascii(text):
    """Printable ASCII whatever a name or a message holds; a runner's console may not be UTF-8."""
    return str(text).encode("ascii", "backslashreplace").decode("ascii")


def _parser():
    parser = argparse.ArgumentParser(
        prog="commit_trailer_check.py",
        description=(
            "Fail when any commit in a range carries an AI co-author trailer, as git itself reads "
            "the commit's trailers. Exit 0 = none, 1 = at least one, 2 = the range could not be read."
        ),
    )
    parser.add_argument("--repo", default=".", help="the clone to read (default: the current directory)")
    parser.add_argument(
        "--base", required=True, help="the push's before SHA, or the pull request's base SHA; all zeros = new branch"
    )
    parser.add_argument("--head", required=True, help="the push's after SHA, or the pull request's head SHA")
    parser.add_argument(
        "--ref", default="", help="the pushed ref (refs/heads/<branch>); required when --base is all zeros"
    )
    parser.add_argument(
        "--forced",
        default="",
        help="the push's forced flag as GitHub renders it, true or false; with true, a --base no clone "
        "holds is read as a new branch instead of failing, unless the push was of --default-branch",
    )
    parser.add_argument(
        "--default-branch",
        default="",
        help="the remote's default branch, github.event.repository.default_branch; a new branch reads its "
        "head minus <remote>/<default-branch>. Required when --base is all zeros or a forced --base is unknown",
    )
    parser.add_argument(
        "--remote", default="origin", help="the remote whose default branch a new branch is read against"
    )
    return parser


def run(argv=None):
    """Judge the range and print the report. Returns the exit code."""
    args = _parser().parse_args(argv)
    repo = args.repo
    out = ["commit-trailer-check: %s" % _ascii(os.path.abspath(repo))]
    try:
        forced = args.forced.strip().lower() == "true"
        commits, desc = read_range(
            repo,
            args.base.strip(),
            args.head.strip(),
            args.ref.strip(),
            args.remote,
            forced,
            args.default_branch.strip(),
        )
        findings = []
        for sha, _subject, trailers in commits:
            hits = trailer_findings(trailers)
            if hits:
                findings.append((sha, _author(repo, sha), hits))
    except RangeError as exc:
        out.append("UNREADABLE -- %s" % _ascii(exc))
        out.append("  No commit message was judged, so this is exit 2 and not a pass.")
        print("\n".join(out))
        return 2
    except Exception as exc:  # an unexpected failure is a range not read, never a verdict
        out.append("UNREADABLE -- unexpected %s: %s" % (type(exc).__name__, _ascii(exc)))
        out.append("  No commit message was judged, so this is exit 2 and not a pass.")
        print("\n".join(out))
        return 2

    out.append("range: %s" % _ascii(desc))
    out.append("%d commit(s) in range" % len(commits))
    if findings:
        for sha, author, hits in findings:
            out.append("AI TRAILER  %s  %s" % (sha, _ascii(author)))
            for line, _value, sig in hits:
                out.append("    %s    [signal: %s]" % (_ascii(line.strip()), ",".join(sig)))
        out.append("FAIL -- %d of %d commit(s) in range carry an AI co-author trailer." % (len(findings), len(commits)))
        out.append("  The trailer is banned on every surface by the decision of 2026-09-18. Reword each")
        out.append("  message above before the branch is merged. A commit already on a shared branch is")
        out.append("  history, and rewriting it is the repository owner's decision, not this check's.")
        print("\n".join(out))
        return 1
    if not commits:
        out.append("PASS (nothing asserted) -- 0 commit(s) in range, so no commit message was judged.")
    else:
        out.append("PASS -- none of the %d commit(s) in range carries an AI co-author trailer." % len(commits))
    print("\n".join(out))
    return 0


if __name__ == "__main__":
    sys.exit(run())
