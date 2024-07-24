## Troubleshooting

> **NOTE**
> Please consider that all is testes using PowerShell 7


### Troubles to execute the script

Because the script is not signed the most common error at the moment can be this one:
![image](https://github.com/user-attachments/assets/52cc87f3-2fdd-4b76-8f6a-c4eafe4324cc)

To resolve this issue you need to execute:
```
Set-ExecutionPolicy -ExecutionPolicy Bypass
```
<br>  

In some organizations this configuration is blocked, like this:
![image](https://github.com/user-attachments/assets/27f17802-1595-49e4-883a-05133ae96c5f)

To resolve this issue you need to execute:
```
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
```
<br>  

**Now, you will be able to execute the script.**

### I can execute the script nevertheless I cannot get info. **PENDING**

If you are executing the script as a Global Administrator, I have to you some news you don't have enough Power, some Microsoft Purview require some additional permissions.
