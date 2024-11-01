# Automated template generator - Generate Karate Templates for APIs from Swagger

This repository contains a tool that generates Karate templates from the swagger.json of an API. This tool can be used with any Swaggger that is version 2 or 3. 
This repo contains the source code of the tool which collects the swagger.json of an API and then using the Swashbuckle class split it into the endpoints.
As well as generating an excel file with an overview of all the endpoints.

## Getting Started

### Prerequisites

    - [.NET SDK](https://dotnet.microsoft.com/download)
    - [Node.js](https://nodejs.org/)

### Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/AutomatedKarateCreation.git
    cd AutomatedKarateCreation
    ```

2. Restore .NET dependencies:
    ```sh
    dotnet restore
    ```

3. Install Node.js dependencies:
    ```sh
    npm install
    ```

### Building the Project

To build the project, run:
    ```sh
    dotnet build
    ```
### Running the project
    ```sh
    dotnet run
    ```

### How to use the tool?

    When running the tool you will be asked for the .json of the swagger. This is commonly found in the swagger itself, either by a link in swagger 3 or in the 
    text bar of swagger 2. Once you provide this, the tool will ask for the authentication of the API, the URL of the API and last but not least the name of the API and folder you want stored.

    Once all of this is done, the tool will generate an excel file under the features folder. It will also create a full and ready to use environment to run karate in using Karate NPM. In order to do so, install Jbang, and then open the terminal in the root of the karate folder and write "npm i" which will install all the required packages.

    Once this is done, it will be possible to run all the karate tests in the template. Initially, they will be mostly filled in with explanations and examples to show you how to fill in the karate tests. 

    In order to run tests, in the terminal you will need to write the command in the following form:

    ```sh
        npm run test -tests=src/test-location  -env=ChosenEnvironment
    ```
    With src/test-location being a placeholder, for you to enter the actual location of the tests you would like to run. And ChosenEnvironment is a placeholder
    for the environment you want to use. For example dev03 or delivery etc.

## License

    This repository is licensed under the MIT License. See the `LICENSE` file for more details.
