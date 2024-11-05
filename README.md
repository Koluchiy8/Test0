This test checks the standard Excel function Today() in Excel Online.

## Steps:
1. Launch the Chrome browser.
2. Navigate to the site excel.com.
3. Authorization (login details are taken from the config.json file).
4. Open a new Excel workbook.
5. Enter the Today() function and read the value.
6. Compare the value with the current date

## Setup
1. Clone the repository.
2. Add your credentials to `config.json`.
3. Installation of necessary libraries for running the test

## Running the Test
Run `npm test` to execute the test in Jest.

## Structure of project
├── src/                  # Source code directory
|    ├──config.json       # A configuration file to store sensitive credentials like email and password.
|    ├──log.txt           # log entries with main operations.
|    ├──test.spec.ts      # The main test file where the end-to-end test is implemented.
│
├── node_modules/         # npm packages
│
├── jest.config.ts        # Jest configuration file
├── package.json          # Project dependencies and scripts
├── tsconfig.json         # TypeScript configuration
├── Demo_video.mkv        # Demo video of the test operation
└── README.md             # Project documentation


## Notes
- To run the test, the following components must be installed: node.js, playwright, jest.
- The test depends on the speed and stability of the internet connection.
- There is a dependency on the website's interface.
- The copy and paste function from the clipboard may work inconsistently; it's better to use a short pause before these operations.
- The format of the date obtained from Excel Online and the built-in function may differ. This conversion is not provided for.
