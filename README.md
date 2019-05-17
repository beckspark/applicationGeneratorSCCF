"# applicationGeneratorSCCF" 

This project is a simple python script meant to automate the administrative case management process of a small nonprofit serving low-income seniors. Rather than spending time redundantly repeating information & data regarding clients and applicants manually in different software for different purposes, this script:

1) Determines for the user whether the applicant qualifies for the programs offered by the organization
2) Generates the client's approval or denial letter
3) Enters the client's data automatically into the dental CMS used by the organization (Note that the CMS is accessed over RDP in Windows, so GUI automation is used in Python to accomplish this).
4) Creates the appropriate internal directories to store the applicant's files for record keeping, and moves the files for the users to those directories.
5) Uses pandas to read & write the applicant's data to the appropriate files used by the organization for future reporting & analysis.

This is one of my first projects in Python, and I am more than open to criticisms & especially suggestions on how to make the code cleaner, more intuitive, & what features I might think of adding in the future.

What I currently plan on changing or adding:

- Accessing and utilizing poverty levels directly from hhs.gov ('https://aspe.hhs.gov/poverty-guidelines'), and using them to automatically determine the approval level for the organization without the need to manually update the numbers year-by-year (the organization approves seniors up to 200% of the federal poverty guidelines).

- Moving the organization's data-keeping files from Dropbox to Google Drive, & updating the appropriate files automatically through the Google Docs & Sheets API.
