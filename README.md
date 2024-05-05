# OneDrive operations via Node.js server

This Node.js application connects to and interacts with OneDrive using Microsoft Graph API. It implements the OAuth 2.0 authorization code flow using the Microsoft Authentication Library (MSAL). Further, it exposes REST APIs  with these functionalities:
* Listing all files in the drive
* Getting all users who can access a given file
* Downloading a file
* Subsrcribe to file's access changes

It also implements a real-time monitoring of file access changes (The file must be subscribed first using the above mentioned API).

The API contracts are documented below. The Postman collection export is shared.

### Prerequisites

Before you begin, ensure you have Node.js and npm (Node Package Manager) installed on your machine. You can install them from [nodejs.org](https://nodejs.org/).

Then, install the project dependencies:
```bash
npm install
```
### Testing purpose, I created a new Microsoft account with these credentials: username: prakhyat2612@outlook.com, password: strac1234567
### You can SKIP to "Execution" section to use the test setup directly.

### Setting up Azure AD Application
* Go to the Azure Portal.
* Navigate to "App registrations" -> "New registration".
* Set the redirect URI (Web) to http://localhost:3000/redirect.
* Go to "API permissions" -> "Add a permission" -> "Microsoft Graph" -> "Application permissions".
* Add necessary permissions - Files.Read.All, Files.ReadWrite.All, Files.ReadWrite.AppFolder, User.Read.
* Navigate to "Authentication".
* Under "Implicit grant and hybrid flows", enable the options for "Access tokens".
* Create a Secret: Go to "Certificates & secrets" -> "New client secret".

### Set the config
* In config.json, set the `clientId` and `clientSecret` from the created Azure AD App.

### Execution
* Run this command:
```bash
node app.js
```
* Open web browser and navigate to http://localhost:3000.
* This will redirect you to the Azure sign-in page. Sign in with the credentials. **Use the credentials (prakhyat2612@outlook.com, strac1234567) for the test setup**.
* After "Login successful!" screen, hit APIs in Postman export given to test functionalties

### API Contracts
#### 1. List All Files
- **Endpoint**: `GET /files`
- **Description**: Retrieves all files in the authenticated user's OneDrive.
- **Sample request**: `http://localhost:3000/files`
- **Sample response**:
    - ```{ "status": 200,"result": {"files": ["testing.txt"]}}```

#### 2. Get Users Accessing a File
- **Endpoint**: `GET /users?file=filename.txt`
- **Description**: Retrieves a list of users who have access to the specified file.
- **Parameters**:
  - `file`: The name of the file (e.g., `testing.txt`)
- **Sample request**: `http://localhost:3000/users?file=testing.txt`
- **Sample response**:
    - ```{ "status": 200,"result": {"users": ["userId1"]}}```

#### 3. Download File
- **Endpoint**: `POST /download/filename.txt`
- **Description**: Downloads the specified file from the user's OneDrive.
- **Parameters**:
  - `file`: The name of the file to download (e.g., `testing.txt`)
- **Sample request**: `http://localhost:3000/download/testing.txt`
- **Sample response**:
    - ```{ "status": 200,"result": "success"}``` 

#### 4. Subscribe to Changes on a File
- **Endpoint**: `POST /subscribe/filename.txt`
- **Description**: Subscribes to notifications for new users accessing the specified file.
- **Parameters**:
  - `file`: The name of the file to monitor (e.g., `testing.txt`)
- **Sample request**: `http://localhost:3000/subscribe/testing.txt`
- **Sample response**:
    - ```{ "status": 200,"result": "success"}``` 
