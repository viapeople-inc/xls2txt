# xls2txt and git: converting spreadsheets to text

This executable when configured with git allows viewing textual diffs of
 spreadsheet files using `git log` or `git diff` rather than getting
 an unhelpful "binary files differ":

## Configure your system

To configure git to work with the executable on your machine, follow these steps:

1. First download the executable from this url: [https://github.com/viapeople-inc/xls2txt/releases/download/v1.2.1/xls2txt.exe](https://github.com/viapeople-inc/xls2txt/releases/download/v1.2.1/xls2txt.exe).

2. Move the executable file to the C:\Deploy Scripts directory.

3. Make sure that your `%PATH%` environment variable contains `C:\Deploy Scripts` so that the executable can be run from the command line.  This can be tested by opening a terminal window and running the command: `xls2txt`.  You should see:

       error: the following required arguments were not provided: <PATH>

        Usage: xls2txt.exe <PATH>

        For more information, try '--help'.

    This shows that the executable is available on the path.

4. Create a `$HOME/.config/git/attributes` file. You may need to create the `.config`
 and/or the `git` directories.

5. In that file, associate the relevant spreadsheet extensions with
   the proper category (hunk-header):

        *.ods diff=spreadsheet
        *.xls diff=spreadsheet
        *.xlsx diff=spreadsheet
        *.xlsb diff=spreadsheet

6. Set `xls2txt`, possibly configured as you desire, as
   the diff text converter:

        git config --global diff.spreadsheet.textconv xls2txt

7. At this point, you are configure!  To test the results, navigate to any directory with an Excel file and make a change to it.  Then run the command: `git diff`.  You will see the change that you just made!

Enjoy!
