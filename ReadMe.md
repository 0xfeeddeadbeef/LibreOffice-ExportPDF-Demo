# How to use LibreOffice in Server Mode to convert Word DOCX to PDF

## Prerequisites

- [LibreOffice](https://www.libreoffice.org/) 64-bit
- .NET Framework >= 4.0


## How to Run

- Installation directory must be added to PATH environment variable
- Create new environment variable UNO_PATH
- Launch LibreOffice in headless server mode (You can use [Non-Sucking Service Manager](http://nssm.cc/) to turn it into a Windows Service)
- **You can use [Launch-LibreOfficeServer.ps1](Launch-LibreOfficeServer.ps1) script**
- Compile and run this program


## Known Issues and Limitations

- Errors thrown by LibreOffice Remoting channel are NOT actionable/useful
- LibreOffice .NET API is based on .NET Remoting (Discontinued by Microsoft, no longer supported)
- Quite often, when error occurs on server-side (for example, when docx file is locked), no error is returned, only local object ends up `null` after casting
- It is slow :stuck_out_tongue:

## Disclaimer

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.

