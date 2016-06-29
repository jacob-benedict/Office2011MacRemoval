#!/bin/bash

###############################################################
#  Office 2011 Mac Removal                                    #
#  officeremoval.sh                                           #
#  v2.0 (6/29/16)                                             #
#  By Jake Benedict                                           #
#  Franklin & Marshall College                                #
#  Reference: https://support.microsoft.com/en-us/kb/2398768  #
###############################################################  


########################
### Aesthetic Set Up ###
########################

printf '\e[8;20;86t' && clear ### Set window size ###

bold=$(tput bold) ### Create Bold Text ###
normal=$(tput sgr0) ### Create Normal Text ###



#######################################
######## Function Declarations ########
#######################################

KillOfficeApplications()
{
	sudo killall Microsoft\ Word &> /dev/null 
	sudo killall Microsoft\ PowerPoint &> /dev/null
	sudo killall Microsoft\ Excel &> /dev/null
	sudo killall Microsoft\ Outlook &> /dev/null 	
}

RemoveOfficeFolder()
{
	echo "Removing the Microsoft Office 2011 Application Folder"
	sudo rm -R /Applications/Microsoft\ Office\ 2011 &> /dev/null
	sudo rm -R /Applications/Microsoft\ Lync.app &> /dev/null
}

RemoveOfficePreferences()
{
	echo "Removing Office Preferences"
	sudo rm /Users/$user/Library/Preferences/com.microsoft* &> /dev/null
}

RemoveOfficeApplicationSupport()
{
	echo "Removing Application Support"
	sudo rm -R /Library/Application Support/Microsoft &> /dev/null
}

RemoveOfficeReceipts()
{
	echo "Removing Office Receipts"
	sudo rm /Library/Receipts/Office2011_* &> /dev/null
	sudo rm /private/var/db/receipts/com.microsoft.office* &> /dev/null
}

ClearMicrosoftCaches()
{
	echo "Clearing Microsoft Caches"
	sudo rm -R /Users/$user/Library/Caches/com.microsoft* &> /dev/null
}

RemoveOfficeLicensingFiles()
{
	echo "Removing Office Licensing Files"
	sudo rm /Library/LaunchDaemons/com.microsoft.office.licensing.helper.plist &> /dev/null
	sudo rm /Library/PrivilegedHelperTools/com.microsoft.office.licensing.helper &> /dev/null
	sudo rm /Library/Preferences/com.microsoft.office.licensing.plist &> /dev/null
	sudo rm /Users/$user/Library/Preferences/ByHost/com.microsoft* &> /dev/null
}

RemoveMicrosoftUserData()
{
	echo "Removing Custom Templates"
	echo "Removing User Data"
	sudo rm -R /Users/$user/Library/Application Support/Microsoft &> /dev/null
	sudo rm -R /Users/$user/Documents/Microsoft User Data &> /dev/null
}

RemoveMicrosoftFonts()
{
	echo "Removing Microsoft fonts"
	sudo rm -R /Library/Fonts/Microsoft &> /dev/null
}

ComputerReboot()
{
	sudo reboot
}

RemovalCancel()
{
	printf "\n\n"; 
	echo "Removal cancelled"
	printf "\n"
	echo "#####################################################################################" 
	printf "\n"
	exit 0
}

#RemoveMicrosoftLicensingData()
##{
#	echo "Would you like to remove the licensing files for Office?"
#	echo "${bold}Note: This could remove licensing for newer versions of Office!${normal}"
#	select yn in "Yes" "No"; do
#	    case $yn in
#			
#			Yes) 
#			echo "Are you ${bold}sure${normal}?";
#			select yn in "Yes" "No"; do
#				case $yn in
#					Yes) 
#					RemoveOfficeLicensingFiles;
#					break;;
#
#					No ) 
#					printf "\n";
#					echo "Licensing files not removed";
#					printf "\n";
#					break;;
#				esac
#			done
#			break;;
#			
#			No ) 
#			printf "\n";
#			break;;
#		esac
#	done
#}

#MicrosoftDataRemoval()
##{
#	echo "Would you like to remove Microsoft User Data?"
#	echo "${bold}Note: This will remove custom user templates!${normal}"
#	select yn in "Yes" "No"; do
#		case $yn in
#
#			Yes)
#			echo "Are you ${bold}sure${normal}?";
#			select yn in "Yes" "No"; do
#				case $yn in
#					Yes) 
#					RemoveMicrosoftUserData;
#					break;;
#
#					No )
#					printf "\n";
#					echo "User Data and Templates not removed.";
#					printf "\n";
#					break;;
#
#			No )
#			printf "\n";
#			break;;
#		esac
#	done
#}

#RemoveMicrosoftFontsQ()
##{
#	echo "Do you want to remove Microsoft fonts?"
#	select yn in "Yes" "No"; do
#	    case $yn in
#			Yes)
#			echo "Are you ${bold}sure${normal}?";
#			select yn in "Yes" "No"; do
#				case $yn in
#					Yes) 
#					RemoveMicrosoftFonts;
#					break;;
#
#					No ) 
#					printf "\n";
#					echo "Microsoft Fonts not removed.";
#					printf "\n";
#					break;;
#			No ) 
#			printf "\n";
#			break;;
#		esac
#	done
#}

########################################
##### End of Function Declarations #####
########################################



#########################
######## Welcome ########
#########################

printf "\n"
echo " ######################## ${bold}Microsoft Office 2011 Removal v2.0${normal} ########################"
printf "\n"
echo "This script will now attempt to completely remove Microsoft Office 2011 from this Mac."
printf "\n"



###############################
### Quit Microsoft Programs ###
###############################

echo "To proceed, we must first quit all Microsoft Apps. You should save any work now." 
printf "\n"
echo "Are you ${bold}sure${normal} you want to continue?"

select yn in "Yes" "No"; do
    case $yn in
		Yes) printf "\n\n"; 
		echo "Please enter your administrator password."; 
		KillOfficeApplications
		break;;
	
		No ) 
		RemovalCancel 
		break;;
	esac
done

printf "\n"
echo "All Microsoft Applications have been quit."
printf "\n"



########################################
### Removal of Microsoft Office 2011 ### 
########################################

echo "Are you ${bold}sure${normal} you want to continue?"
select yn in "Yes" "No"; do
    case $yn in
		
		Yes) 
		printf "\n \n";
		echo This is happening.;
		printf "\n";
		
		RemoveOfficeFolder;
		RemoveOfficePreferences;
		ClearMicrosoftCaches;
		RemoveOfficeApplicationSupport;
		RemoveOfficeReceipts;
		
		printf "\n";
		echo " ###########################################################################";
		printf "\n";
		
		echo "Remove all Microsoft icons from the dock."
		echo "Once you have done so, press any key to continue... ";
		
		read -n 1;
		echo;
		break;;

		No )
		RemovalCancel;
		break;;
	esac
done

###############################
##### End of Main Removal #####
###############################



####################################
###### Advanced Removal Tools ######
####################################

#echo "Would you like to access the Advanced Removal Tools for Licensing, User Data, and Fonts?"
#select yn in "Yes" "No"; do
#    case $yn in
#		Yes) 
#		printf "\n \n";
#		RemoveMicrosoftLicensingData;
#		MicrosoftDataRemoval;
#		RemoveMicrosoftFontsQ
#		break;;
#
#		No )
#		printf "\n \n";
#		break;;
#	esac
#done



##########################
### Reboot and Goodbye ###
##########################

echo " #####################################################################################"
printf "\n"
echo "Thank you for using the Microsoft Office 2011 removal script!"

echo "You should now restart the machine to complete the process."
echo "Would you like to restart now?"
select yn in "Yes" "No"; do
    case $yn in
		Yes) 
		printf "\n \n";
		ComputerReboot;
		exit 1;
		break;;

		No )
		printf "\n \n";
		echo "Please restart at your earliest convenience.";
		break;;
	esac
done
printf "\n"