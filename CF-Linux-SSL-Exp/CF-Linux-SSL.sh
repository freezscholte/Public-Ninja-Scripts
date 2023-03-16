#!/bin/bash

# Set the directory containing the SSL certificates
SSL_CERT_DIR="/etc/apache2"

# Set the number of days to check for certificate expiration
DAYS_THRESHOLD=21

# Initialize the CustomField variable
CustomField=""

# Iterate through the SSL certificates in the specified directory
for cert in $(find "$SSL_CERT_DIR" -type f -name "*.crt" -o -name "*.pem"); do
  # Get the certificate's expiration date in seconds since epoch
  cert_expiration_date=$(openssl x509 -in "$cert" -enddate -noout | cut -d= -f2 | xargs -I {} date -d {} +%s)

  # Get the current date in seconds since epoch
  current_date=$(date +%s)

  # Calculate the remaining days until the certificate expires
  remaining_days=$(( ($cert_expiration_date - $current_date) / 86400 ))

  # Check if the certificate will expire within the next DAYS_THRESHOLD days
  if [ "$remaining_days" -le "$DAYS_THRESHOLD" ] && [ "$remaining_days" -gt 0 ]; then
  
    CustomField="$CustomField Certificate $cert will expire within the next $remaining_days days and is active on Apache or Nginx."
    /opt/NinjaRMMAgent/programdata/ninjarmm-cli set sslCertificates "$CustomField"
  fi
done

# Output the CustomField variable
echo "$CustomField"


# If CustomField is still empty, set it to "All good no expiring certificates found"
if [ -z "$CustomField" ]; then
    CustomField="All good no expiring certificates found"
    /opt/NinjaRMMAgent/programdata/ninjarmm-cli set sslCertificates "$CustomField"
fi




# Print the contents of the CustomField variable
echo "$CustomField"

exit
