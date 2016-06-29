#!/bin/bash

###### OfficeRemoval.sh #####################################
# by Jake Benedict                                          #
# v1.5 (05/11/16)                                           #
# Reference: https://support.microsoft.com/en-us/kb/2398768 #
#############################################################


### Aesthetic Set Up ###

printf '\e[8;20;86t' && clear ### Set window size ###

bold=$(tput bold) ### Create Bold Text ###
normal=$(tput sgr0) ### Create Normal Text ###

### Welcome and Quit Microsoft Programs ###

printf "\n"
echo " ######################## ${bold}Microsoft Office 2011 Removal v1.5${normal} ########################"
printf "\n"
echo "This script will now attempt to completely remove Microsoft Office 2011 from this Mac."
printf "\n"
echo "To proceed, we must first quit all Microsoft Apps. You should save any work now." 
printf "\n"
echo "Are you ${bold}sure${normal} you want to continue?"
select yn in "Yes" "No"; do
    case $yn in
		Yes) printf "\n\n"; 
		echo "Please enter your administrator password."; 
		sudo killall Microsoft\ Word &> /dev/null; 
		sudo killall Microsoft\ PowerPoint &> /dev/null; 
		sudo killall Microsoft\ Excel &> /dev/null; 
		sudo killall Microsoft\ Outlook &> /dev/null; 
		break;;
	
		No ) printf "\n\n"; 
		echo "Removal cancelled"; 
		printf "\n"; 
		echo "#####################################################################################"; 
		printf "\n"; 
		exit 0; 
		break;;
	esac
done
printf "\n"
echo "All Microsoft Applications have been quit."
printf "\n"

### Reomval of Microsoft Office 2011 ### 

echo "Are you ${bold}sure${normal} you want to continue?"
select yn in "Yes" "No"; do
    case $yn in
		Yes) printf "\n \n";
		echo This is happening.;
		printf "\n";
		echo "Removing the Microsoft Office 2011 Application Folder";
		sudo rm -R /Applications/Microsoft\ Office\ 2011 &> /dev/null;
		sudo rm /Applications/Microsoft\ Lync.app &> /dev/null;
		echo "Removing Office Preferences";
		sudo rm /Users/itsadmin/Library/Preferences/com.microsoft* &> /dev/null;
		echo "Removing Office Licensing Files";
		sudo rm /Library/LaunchDaemons/com.microsoft.office.licensing.helper.plist &> /dev/null;
		sudo rm /Library/PrivilegedHelperTools/com.microsoft.office.licensing.helper &> /dev/null;
		sudo rm /Library/Preferences/com.microsoft.office.licensing.plist &> /dev/null;
		sudo rm /Users/itsadmin/Library/Preferences/ByHost/com.microsoft* &> /dev/null;
		echo "Clearing Microsoft Caches";
		sudo rm -R /Users/itsadmin/Library/Caches/com.microsoft* &> /dev/null;
		echo "Removing Application Support";
		sudo rm -R /Library/Application Support/Microsoft &> /dev/null;
		echo "Removing Receipts";
		sudo rm /Library/Receipts/Office2011_* &> /dev/null;
		sudo rm /private/var/db/receipts/com.microsoft.office* &> /dev/null;
		echo "Removing Custom Templates";
		sudo rm -R /Users/itsadmin/Library/Application Support/Microsoft &> /dev/null;
		echo "Removing User Data";
		sudo rm -R /Users/itsadmin/Documents/Microsoft User Data &> /dev/null;
		printf "\n";
		echo " ###########################################################################";
		printf "\n";
		echo "Remove all Microsoft icons from the dock."
		echo "Once you have done so, press any key to continue... ";
		read -n 1;
		echo;
		break;;

		No )
		printf "\n\n";
		echo "Removal cancelled";
		printf "\n";
		echo " #####################################################################################";
		printf "\n";
		exit 0;
		break;;
	esac
done

### End of Removal ###

echo " #####################################################################################"
printf "\n"
echo "Thank you for using the Microsoft Office 2011 removal script!"

echo "You should now restart the machine to complete the process."
echo "Would you like to restart now?"
select yn in "Yes" "No"; do
    case $yn in
		Yes) printf “\n \n”;
		sudo reboot;
		exit 1;
		break;;

		No )
		printf "\n \n";
		echo "Please restart at your earliest convenience.";
		break;;
	esac
done
printf "\n"
