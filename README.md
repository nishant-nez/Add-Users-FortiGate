# Add-Users-FortiGate

A python automation code for deleting usernames from FortiGate 100F firewall using selenium built for DevOps department, DWIT.

## Contents

- **main.py**: The main python file that runs the selenium code.
- **credentials.ini**: A ini file containing the username, password and IP address to open the firewall page.

## Installation

Selenium package is required to run this Program.

Install the dependencies if not exists

```sh
pip install selenium
```

The following modules are also required

Pandas:

```sh
pip install pandas
```

Configparser:

```sh
pip install configparser
```

OS:

```sh
pip install os
```

Random:

```sh
pip install random
```

---

## How to Run

- Open credentials.ini and replace the values of link, username and password with correct value.
- Run main.py and follow the prompts.
