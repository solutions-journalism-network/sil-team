## Our team website

The database is a Google Sheet with two tabs: People and Contracts. This allows us to add zero, one, or multiple contracts per person.

Admin can edit the data using the website OR edit the spreadsheet directly. If adding a new people or contract record, use the website so it will auto-generate the ID numbers.

The website never shows or loads emails and phone numbers. Users must go to the spreadsheet to locate this information.


## How status works

Each person's status (the row colour on the site) is worked out from their contracts, so it never goes stale:

| Colour | Status | When |
|---|---|---|
| Teal | **On contract** | they have a contract whose dates include today (no end date counts as ongoing) |
| Amber | **Getting on contract** | no active contract, but one that starts in the future |
| Blue-grey | **Prospect** | no contracts yet |
| Grey | **Past** | they have contracts but none are active |

For a renewal, add a *new* contract row with the new dates. The old one stays as history, and shows as Ended.

## Daily use

- **Add a contractor or a contract:** use **+ Add contractor** / **+ Add contract** on the site, or type into the Sheet. On *Contracts*, pick the person in the *Name* column (it's a dropdown). IDs fill in by themselves: don't edit them.
- **Edits made by hand in the Sheet** show up on the site right away. If you use File > Import or a script to change lots of rows, run Contractors > Refresh the website's data.
- **Two people editing at once:** if someone saves a person or contract that another person changed in the meantime, the site refuses the second save, shows the latest version, and asks them to try again. Nothing is overwritten silently. If two saves collide at the same moment, the site says it's busy: try again in a moment.
- **Check data** (Contractors menu): lists contracts with no matching person, people listed twice, people with no role or country, end dates before start dates, and so on.
- **Deleting a person** also deletes their contracts. Undo with File > Version history in the Sheet.
- **Change the password:** Script Properties. Do this when someone with the password leaves the team. No redeploy needed.
- **After editing Code.gs:** Deploy > Manage deployments > pencil icon > Version: New version > Deploy. The URL stays the same.
- **Columns:** you can reorder columns and add your own on either tab. The script finds columns by their headings, so don't rename these: Person ID, Name, Role, City/State, Country, Language, Bio, CV link, Email, Phone, Notes (People) and Contract ID, Name, Projects, Start date, End date, Payment terms, Notes, Person ID (Contracts). It only ever writes to those columns, so your own columns and formulas are left alone.
- **Lists in a cell** (roles, languages, projects) are separated with semicolons or commas. The site writes them as `Interviewer; Editor`.
- **Dates** are real dates formatted `yyyy-mm-dd`. The site writes them the same way.
- The role, country, language and project lists on the site (filters and quick-pick buttons) come from what's already in the Sheet, so keep spellings consistent (for example "United States", not "USA").

## Good to know
- After 10 wrong passwords, everyone is locked out for 10 minutes.
- The "Ends in N days" flag appears for people on contract within 60 days of their contract's end date (`ENDING_SOON_DAYS` in `index.html`). If someone has several active contracts, it uses the last one to end.
- The site is hidden from search engines (`noindex`), but the password is what protects the data.
