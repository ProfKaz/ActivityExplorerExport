## Troubleshooting

> **NOTE**
> Please consider that all is testes using PowerShell 7

Because the script is not signed the most common error at the moment can be this one:
![image](https://github.com/user-attachments/assets/52cc87f3-2fdd-4b76-8f6a-c4eafe4324cc)

To resolve this issue you need to execute:
```
Set-ExecutionPolicy -ExecutionPolicy Bypass
```

In some organizations this configuration is blocked, like this:
![image](https://github.com/user-attachments/assets/27f17802-1595-49e4-883a-05133ae96c5f)


To resolve this issue you need to execute:
```
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
```

**Now, you will be able to execute the script.**
