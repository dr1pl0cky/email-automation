## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### Prerequisites

What things you need to install the software and how to install them. Run as admin in powershell

```
Install-Module PSExcel
```

### Installing

Set the execution policy so it will run, do this also as admin on powershell. 

```
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine
```

## Running the tests

Using the dummy data in `Book1.xlsx` and files such as `test1` etc. You can test the script to make sure its working before using customer data. Make sure the xlsx has said filenames correct for each user.

Change the email and password in script as needed and make sure you go to your email providor and enable 3rd party apps such as gmail.

Gmail example:
`Account Settings > Security > Less secure apps > ON`

Change code for parameters that match xlsx headers. Running will spit out an error that 1st line in output.csv is not a valid email and is just because headers are in place. They are needed however, this will be fixed in another version, probably. Maybe. 


## Deployment

Run the powershell script in ISE or powershell terminal

`email-auto.ps1`


## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE.md](LICENSE.md) file for details

