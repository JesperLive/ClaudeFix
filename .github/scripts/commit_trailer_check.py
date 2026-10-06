#!/usr/bin/env python3
r"""commit_trailer_check.py -- fail a commit range that carries an AI co-author trailer.

Created: 2026-10-06
Updated: 2026-10-06 (a "(cherry picked from commit" line continues the final trailer block, as git's own trailer parser counts it, so a model trailer that git cherry-pick -x carries over is read; earlier the same day prompt 10 of the 2026-09-28 tools queue: new; the trailer pattern moved here verbatim from cowork_util.py's audit-commit-trailer section, so that gate and every repo's CI judge a message with one copy of it)

Usage, as every repo's .github/workflows/commit-trailer.yml runs it:

    python3 .github/scripts/commit_trailer_check.py --repo . --base <before> --head <after> --ref <ref> --forced <true|false>
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

ONE FILE, TWO READERS. The source lives at tools/hooks/commit_trailer_check.py
in GRIP-Tools. cowork_util.py loads it from there and audit-commit-trailer judges
every commit through its functions, so the gate that reads the whole workspace
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
    is every commit reachable from the pushed SHA and from no other branch of
    the remote. The runner's clone holds the pushed branch's own remote ref and
    may hold the remote's HEAD symref, so both are left out of "other";
  * a FORCED push whose before SHA no clone holds -- a rebase pushed with
    --force-with-lease, as Dependabot does to its own branches -- takes the
    new-branch rule, because GitHub marks the push forced (--forced true); the
    2026-10-06 review measured every such push failing at exit 2 without it;
  * a pull request hands its base and head SHAs, and the range is base..head.
A range needs the history under it, so the workflow checks out with
fetch-depth 0; actions/checkout@v7 defaults to 1, measured in the runner's cached
copy of the action, and a depth-1 clone cannot resolve a before SHA at all. The
new-branch rule cannot see a second branch created by the same push; see
_branch_revisions for that limit and where those commits are read instead.
"""

import argparse
import os
import re
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
AI_VENDORS = ("openai", "chatgpt", "copilot", "gemini", "codex", "devin", "aider")

# A trailer key: a letter, then letters, digits or hyphens, then a colon, then a
# non-empty value. The optional [ \t]* before \S is what admits the real shape --
# `Co-Authored-By: Name <addr>` puts a space after the colon, and a pattern that
# demanded a non-space IMMEDIATELY after it would match none of the 82 commits
# this gate exists for. What the \S still buys is that a bare `Fixes:` with
# nothing after it is not a trailer, so it terminates the block rather than
# extending it.
KEY_RE = re.compile(r"^[A-Za-z][A-Za-z0-9-]*:[ \t]*\S")

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
ATTRIBUTION_KEYS = ("author", "co-author", "coauthor", "cc")

# A LINE GIT WRITES INTO THE TRAILER BLOCK WITH NO KEY, measured 2026-10-06 on
# git 2.55.0.windows.3 rather than read off the documentation. `git cherry-pick
# -x` appends "(cherry picked from commit <sha>)" straight under a message's
# trailer block, and `-x -s` appends a Signed-off-by under that. git's own trailer
# parser counts the line as part of the block -- it is one of git's
# "git-generated" prefixes -- and `git interpret-trailers --parse` still returns
# a Co-Authored-By above it. A block parse that stopped at the first line not
# shaped "Key: value" read such a message's block as empty, or as the sign-off
# alone, so a cherry-picked model Co-Authored-By passed at exit 0, and the
# commit-msg hook does not run on a cherry-pick at all. The line now continues the
# block and is never itself an attribution, having no key. A commit with no
# trailer block gets the line after a blank line instead, so nothing above it is
# read. Scored the same day over the full history of the 14 owned repos, 3475
# commits: no message carries the line, and the 86 AI attribution lines and 93
# signal-bearing block lines read identically under both rules, so this closes the
# hole and moves no verdict.
CHERRY_PICK_PREFIX = "(cherry picked from commit "


def is_attribution(line):
    """True when `line`'s trailer KEY attributes authorship to somebody."""
    key = line.partition(":")[0].strip().lower()
    return key.endswith(ATTRIBUTION_KEY_SUFFIXES) or key in ATTRIBUTION_KEYS


def final_block(message):
    r"""The message's FINAL TRAILER BLOCK: the last contiguous run of trailer lines.

    Terminated upward by the first blank line or the first line that is not
    trailer-shaped. Returns [] when the last non-blank line is not itself
    trailer-shaped -- which is the case for every ordinary commit message, and
    the reason a body that merely MENTIONS Claude or CLAUDE.md is out of range
    without any exception list. A line git's cherry-pick -x writes, which starts
    with CHERRY_PICK_PREFIX, continues the block as git's own parser lets it.

    Trailing blank lines are dropped first, because `git log --format=%B` ends
    every body with one and a naive "last line" read would see it and return
    empty for every commit in the repository.
    """
    lines = (message or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    while lines and not lines[-1].strip():
        lines.pop()
    block = []
    for line in reversed(lines):
        if not line.strip():
            break
        if not KEY_RE.match(line) and not line.startswith(CHERRY_PICK_PREFIX):
            break
        block.append(line)
    block.reverse()
    return block


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


def ai_attribution_lines(message):
    """[(line, value, hits)] for each final-block attribution line whose value carries a signal.

    THE WHOLE JUDGEMENT OF ONE MESSAGE, in one function both readers call:
    audit-commit-trailer applies its declared exceptions to what this returns,
    and the CI check fails on anything it returns. A message whose list is
    empty carries no AI co-author trailer by this file's definition.
    """
    found = []
    for line in final_block(message):
        if not is_attribution(line):
            continue
        value = trailer_value(line)
        hits = signals(value)
        if hits:
            found.append((line, value, hits))
    return found


def parse_log(payload):
    """Split `%H%x1f%B%x1e` output into (sha, body) pairs.

    THIS EXACT SHAPE WAS PROVEN AGAINST ALL 11 REPOS ON 2026-09-16. The record
    separator is what makes a multi-line body unambiguous -- a body can contain
    blank lines, indented text and its own colons, so no line-oriented split
    survives contact with a real commit message.
    """
    commits = []
    for record in payload.split("\x1e"):
        record = record.lstrip("\n")
        if not record.strip():
            continue
        sha, sep, body = record.partition("\x1f")
        if not sep:
            continue
        commits.append((sha.strip(), body))
    return commits


# ------------------------------------------------------------------
# THE RANGE. Everything below reads git and nothing above does.
# ------------------------------------------------------------------

LOG_FORMAT = "--format=%H%x1f%B%x1e"


class RangeError(Exception):
    """The range could not be read. main() turns it into exit 2, never 0."""


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


def range_revisions(repo, base, head, ref="", remote="origin", forced=False):
    """(git log revision arguments, one-line description) for the range, or RangeError.

    base..head when base is a commit. When base is all zeros the push created
    a branch, and the range is the new-branch rule's, _branch_revisions. A
    FORCED push whose base is a commit this clone does not hold -- the old tip
    of a rebase, reachable from nothing once the push lands -- takes the same
    rule; an unforced one with an unknown base stays unreadable.
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
        return _branch_revisions(repo, head_full, ref, remote, "the base SHA is all zeros, which names a new branch")
    try:
        base_full = _resolve(repo, base, "base")
    except RangeError:
        if not forced:
            raise
        return _branch_revisions(
            repo, head_full, ref, remote, "the push was forced and its old tip %s is on no branch of this clone" % base
        )
    return ["%s..%s" % (base_full, head_full)], "%s..%s" % (base_full, head_full)


def _branch_revisions(repo, head_full, ref, remote, why):
    """The new-branch rule: head minus every branch of `remote` but the pushed one and the HEAD symref.

    Those two refs point at the pushed branch itself in the runner's clone, so
    counting either as "another branch" would leave nothing to read. THE ONE
    THING IT CANNOT SEE is another branch created by the same push: two new
    branches pushed together at one commit hide each other's commits (found by
    the 2026-10-06 review, reproduced with `git push origin x:refs/heads/a
    x:refs/heads/b`). Those commits are still read the day they reach a branch
    that existed, as a push's before..after or a pull request's base..head.
    """
    if not ref:
        raise RangeError(
            "%s, and no --ref says which branch was pushed -- without it the branch's own remote ref "
            "would count as another branch" % why
        )
    if ref.startswith("refs/heads/"):
        branch = ref[len("refs/heads/") :]
    elif ref.startswith("refs/"):
        branch = None
    else:
        branch = ref
    rc, out, err = _git(repo, "for-each-ref", "--format=%(refname)", "refs/remotes/" + remote)
    if rc != 0:
        raise RangeError("git for-each-ref failed reading the branches of %s: %s" % (remote, err.strip()))
    own = {"refs/remotes/%s/HEAD" % remote}
    if branch:
        own.add("refs/remotes/%s/%s" % (remote, branch))
    others = [r for r in out.split() if r not in own]
    desc = "%s, read as a new branch (%s): commits reachable from %s and from none of the %d other branch(es) of %s" % (
        ref,
        why,
        head_full,
        len(others),
        remote,
    )
    return [head_full, "--not", *others], desc


def read_range(repo, base, head, ref="", remote="origin", forced=False):
    """([(sha, body)], description) for every commit in the range, or RangeError."""
    revisions, desc = range_revisions(repo, base, head, ref, remote, forced)
    rc, payload, err = _git(repo, "log", LOG_FORMAT, *revisions)
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
            "Fail when any commit in a range carries an AI co-author trailer in its final trailer "
            "block. Exit 0 = none, 1 = at least one, 2 = the range could not be read."
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
        "holds is read as a new branch instead of failing",
    )
    parser.add_argument("--remote", default="origin", help="the remote whose branches a new branch is measured against")
    return parser


def run(argv=None):
    """Judge the range and print the report. Returns the exit code."""
    args = _parser().parse_args(argv)
    repo = args.repo
    out = ["commit-trailer-check: %s" % _ascii(os.path.abspath(repo))]
    try:
        forced = args.forced.strip().lower() == "true"
        commits, desc = read_range(repo, args.base.strip(), args.head.strip(), args.ref.strip(), args.remote, forced)
        findings = []
        for sha, body in commits:
            hits = ai_attribution_lines(body)
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
