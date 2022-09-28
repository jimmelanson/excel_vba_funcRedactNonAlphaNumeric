# excel_vba_funcRedactNonAlphaNumeric

funcRedactNonAlphaNumeric
Returns a boolean TRUE or FALSE

===============================================================================

Sometimes you want to have user enter data but you do not want them entering punctuation.
Typically, this would be so that you can process the input with a minimum amount of time
spent on cleaning up the data.

This procedure will remove all non alphanumeric characters (like punctuation) but the
original procedure will still allow user entered spaces to be maintained.

For example, "It's a NP 3607." becomes "Its a NP3607"

I have also included two additional versions of the procedure (commented out) that you
can use for these results:

Alternate No. 1: "It's a NP-3607." becomes "ItsaNP3607"
Alternate No. 2: "It's a NP 3607." becomes "It s a NP 3607"

The *.bas module contains a subroutine to test the procedure.

To use this code, copy and paste it into your project OR import the file worksheet_exists.bas
