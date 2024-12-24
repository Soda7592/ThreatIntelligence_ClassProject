## Project introdution
This project derivatived from [NSFOCUS](https://nsfocusglobal.com/apt34-unleashes-new-wave-of-phishing-attack-with-variant-of-sidetwist-trojan/)

We simulation it and adjust its entry point to a malicious email with a office attachment.

This attachment include the `macro.vbs` and then you need to deploy a http server (you can make easy by python http server) and modify the file name to fit the situation finally the `macro.vbs` can get the malware.

This http shell from [HTTP-Shell](https://github.com/JoelGMSec/HTTP-Shell/tree/main)
