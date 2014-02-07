Out-TFSReportsToCSV
===================

NAME
    out-TFSProjectsToCSV
    
SYNOPSIS
    Parse the output from TFSProjects utility and output as a CSV file
    
    
SYNTAX
    out-TFSProjectsToCSV [[-TFSProjectsFile] <Object>] [[-CSVOutputFileFolder] <Object>] [<CommonParameters>]
    
    
DESCRIPTION
    Take the RTF file generated from TFSProjects (https://tfsprojects.codeplex.com/)
    and convert it to a usable CSV file
    
    ##############################################################################################
    ##  This script requires the free Quest ActiveRoles Management Shell for Active Directory
    ##  snapin  http://www.quest.com/powershell/activeroles-server.aspx
    ##############################################################################################
    

PARAMETERS
    -TFSProjectsFile <Object>
        (Optional) The name of the RTF file that you saved from TFSProjects.
        
        If not speficied, you will be presented with a file-dialog to choose the input file
        
    -CSVOutputFileFolder <Object>
        (Optional) The folder where the CSV output will be place.  If not specified,
        you will be presented with a folder-dialog to choose the destination
        
    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see 
        about_CommonParameters (http://go.microsoft.com/fwlink/?LinkID=113216). 
    
    -------------------------- EXAMPLE 1 --------------------------
    
    C:\PS>Out-TFSProjectsToCSV -TFSProjectsFile tfsp.rtf -CSVOutputFileFolder c:\temp
    
    This will take the tsfp.rtf file, created by TFSProjects, and output
    a CSV file called tsfp.csv to the c:\temp folder
    
    
    
    
    -------------------------- EXAMPLE 2 --------------------------
    
    C:\PS>Out-TFSProjectsToCSV
    
    Without specifying any options, you will be presented with a file-dialog to
    choose the input file (RTF), and a folder-dialog to choose the destination
    for the CSV file
    
    
    
    
REMARKS
    To see the examples, type: "get-help out-TFSProjectsToCSV -examples".
    For more information, type: "get-help out-TFSProjectsToCSV -detailed".
    For technical information, type: "get-help out-TFSProjectsToCSV -full".
    For online help, type: "get-help out-TFSProjectsToCSV -online"




