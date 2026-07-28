@{
    # PSScriptAnalyzer configuration for the ClaudeFix toolkit.
    #
    # Every exclusion below is justified. The point of excluding anything at
    # all is so that a finding in CI means something and gets acted on, rather
    # than being one more line in a wall of noise nobody reads.

    Severity = @('Error', 'Warning')

    ExcludeRules = @(
        # These are interactive console tools. Write-Host is the correct way to
        # put coloured status in front of a person who double-clicked a .bat
        # file, and the output is deliberately not part of a pipeline. 159
        # occurrences at the time of writing, none of them a defect.
        'PSAvoidUsingWriteHost',

        # False positive, verified by running the pattern. Every flagged
        # scriptblock declares param(...) and is fed via -ArgumentList, which
        # is the correct mechanism for passing values into a Start-Job
        # runspace. Using $using: in those scriptblocks would be wrong and
        # would fail at run time.
        'PSUseUsingScopeModifierInNewRunspaces',

        # Cosmetic, and following it would make the names worse.
        # Close-StaleHcsVms closes several VMs; Test-LogsForErrors scans
        # several logs. Renaming them to singular would misdescribe what they
        # do in order to satisfy a naming convention.
        'PSUseSingularNouns'
    )

    # PSAvoidUsingEmptyCatchBlock is deliberately NOT excluded.
    #
    # There were 173 empty catch blocks across the three scripts. All of them
    # now carry a statement and a reason. Keeping the rule active means the
    # next one added without a reason fails the build, which matters: a silent
    # catch is how the temp-file cleanup came to never run at all, and how a
    # user-activity guard came to answer "nobody is here" when it had actually
    # failed.
}
