Using certutil for Enrollment
You can also use certutil to trigger the certificate enrollment. This method is generally used for Windows environments where you want to request a certificate and immediately install it.

powershell
Copy
# Request the user certificate via certutil
certutil -pulse

# To trigger certificate enrollment, use certutil to request from the CA
certutil -renew -user