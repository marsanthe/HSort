
using namespace System.Collections.Generic

$ErrorActionPreference = "Stop"
#$DebugPreference = "SilentlyContinue"
$DebugPreference = "Continue"

function Test-NodeType{
    <# 
    .DESCRIPTION
    Returns 0 if a path points to a leaf folder or 1 if not.
    #>
    Param(
        [Parameter(Mandatory)]
        [Object]$Node
    )
    
    $IsLeaf = 999
    $ChildPath = $Node.FullName
    $DirCount = (Get-ChildItem -LiteralPath $ChildPath -Directory).Count

    If($DirCount -gt 0){
        $IsLeaf = 1
    }
    else{
        $IsLeaf = 0
    }

    return $IsLeaf
}
function Export-SrcDirTree {
    <#
    .SYNOPSIS
    Writes to a .txt file.

    .DESCRIPTION
    Converts the (Key,Value) pairs of a hashtable [SrcDirTree] to a string and writes it to a .txt file.
    [SrcDirTree] is an ordered hashtable representing a modified version of the direcory tree of a source folder
    relative to itself (source folder as root).
    It is modified in so far, as it does not contain CD-folders.

    .NOTES
    When exporting as .txt make sure that there are no unwanted line breaks to avoid parsing errors.
    #>
    Param(

        [Parameter(Mandatory)]
        [string]$LibraryName,

        [Parameter(Mandatory)]
        [string]$SourceDir,

        [Parameter(Mandatory)]
        [System.Collections.Specialized.OrderedDictionary]$SrcDirTree
    )

    $TargetPath = "$($PathsProgram.LibFiles)\$LibraryName\SrcDirTree $LibraryName.txt"

    # Write the head of the section to the txt file
    "[$SourceDir]" | Out-File -LiteralPath $TargetPath -Encoding unicode -Append -Force

     foreach ($Pnode in $SrcDirTree.Keys) {
        
        Write-Debug "GetObjects: Parent: $Pnode"
        
        "Pnode = $Pnode" | Out-File -LiteralPath $TargetPath -Encoding unicode -Append -Force

        # List of direct child-nodes
        $Cnodes = $SrcDirTree.$Pnode

        # If list is not empty...
        if ($Cnodes.Count -gt 0) {

            # ... add child-node entry
            foreach ($Cnode in $Cnodes) {

                Write-Debug "GetObjects: Child: $Cnode"

                "Cnode = $Cnode" | Out-File -LiteralPath $TargetPath -Encoding unicode -Append -Force

            }
        }
    }

    " " | Out-File -LiteralPath $TargetPath -Encoding unicode -Append -Force
}

function Get-Files{
    <# 
    .SYNOPSIS
    Adds an object to a list or 
    adds a hashtable to another hashtable.

    .DESCRIPTION
    Scans the files of a folder.
    If a file matches certain criteria, it is added to [SupportedObjects] list.
    If it doesn't, its properties are stored in the [ExcludedObjects] hashtable.
        
        ExcludedObjects = @{
            ...
            FilePath = @{FileProperties}
            ...
        }
    #>
    Param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [list[object]]$SupportedObjects,

        [Parameter(Mandatory)]
        [hashtable]$ExcludedObjects,

        [Parameter(Mandatory)]
        [hashtable]$VisitedObjects,

        [Parameter(Mandatory)]
        [array]$Counter
    )

    $LooseFiles = (Get-ChildItem -LiteralPath $NodePath -File) # <- File (!)
    
    # Check if current node contains files. 
    if(($LooseFiles.Count -gt 0)){

        Write-Debug "GetObjects: [$($NodePath.Name)] contains files."

        foreach($LooseFile in $LooseFiles){

            $LooseFilePath = $LooseFile.FullName

            if (  ! $VisitedObjects.ContainsKey($LooseFilePath)  ) {

                $null = $VisitedObjects.Add($LooseFilePath, "0")

                $LooseFileExtension = $LooseFile.Extension
                Write-Debug "GetObjects: [Loose file extension] $LooseFileExtension"

                # If [LooseFile] in current node is a supported archive,
                # add it to [SupportedObjects].
                if ( $AllowedArchives.Contains($LooseFileExtension) ) {
                    Write-Debug "GetObjects: File is supported."
                    $null = $SupportedObjects.Add($LooseFile)
                }
                # Else, add it to ExcludedObjects.
                else {
                    Write-Debug "GetObjects: File is NOT supported."
                    $Counter[0]++

                    $ParentPath         = Split-Path -Parent $LooseFilePath
                    $RelativePath       = Get-RelativePath -Path $LooseFilePath -StartDir $SrcDirName
                    $RelativeParentPath = Get-RelativePath -Path $ParentPath -StartDir $SrcDirName
                    
                    # These properties are needed for MoveExcludedObjects.ps1 in Tools
                    $ExcludedObjectProperties = @{
                        Path         = $LooseFilePath;
                        RelativePath = $RelativePath;
                        ParentPath   = $ParentPath;
                        RelativeParentPath = $RelativeParentPath;
                        ObjectName   = $LooseFile.Name;
                        Reason       = "UnsupportedFile";
                        Extension    = $LooseFileExtension
                    }
                    
                    $null = $ExcludedObjects.Add($LooseFilePath, $ExcludedObjectProperties)
                    
                }
            }
        }
    }
    else{
        Write-Debug "GetObjects: GetObjects: [Current node] contains no files"
    }
}

function Get-RelativePath {
    <# 
    .DESCRIPTION
    Returns the path of an object relative to
    a provided direcory name.
    #>
    Param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$StartDir
    )

    $Tail = $Path -Split ($StartDir)
    $RelativePath = $StartDir + $Tail[1]

    if (-not $RelativePath) {
        throw "RelativePath is null or empty string"
    }

    return $RelativePath

}

function Initialize-Graph{
    <#
    .SYNOPSIS
    Creates a hashtable from a root directory
    that contains all directories and their direct children.

    .INPUTS
    A source directory [SrcDir], also called "root", as string.

    .DESCRIPTION
    Creates a hashtable called [Graph] from a root directory.
    [Graph] contains all directories and their direct children in root,
    with the exception of leaf nodes

    This is a hashtable of edges of the form
    
    [Graph]
    Hashtable layout:

        Graph = @{
            [string]DirectoryPath = [Array[System.IO.DirectoryInfo]]ChildDirs;
            .
            .
            .

        }

    #>
    Param(
        [Parameter(Mandatory)]
        [string]$SourceDir
    )

    Write-Output "Scanning folders. Please wait..."

    $Graph = @{}

    # Get all nodes (folders) at root, including those directly at root - recursively.
    # This array does _not_ contain the root folder iteself, but every other
    # directory, including leaf nodes.
    $AllNodes_ObjArr = (Get-ChildItem -LiteralPath $SourceDir -Directory -Recurse)

    # Allows us to access leaf nodes as directory objects
    # Needs to be iterable and extendable
    $LeafNodes_Lst = [List[object]]::new()
        
    # Helps us to quickly check if a node is a leaf node.
    # Needs to be extendable and have a short access time 
    $LeafNodes_HashTable = @{}
    # Update 11/07/2024
    # This was initially a HashSet.
    # $LeafNodes_Set = [HashSet[string]]::new()
    # The HashSet was replaced by a HashTable since it created strange errors.
    # Albeit a string ($SourceDir) was not an element of the HashSet,
    # HashSet.Contains($SourceDir) always evaluated to $true

    #===========================================================================
    # Insert Root-node
    #===========================================================================

    Write-Debug "GetObjects: [Inserting root node]" 
    
    # Get all folders located at root (Direct children of Root).
    $RootChildDirs = (Get-ChildItem -LiteralPath $SourceDir -Directory)

    # If root itself is not a leaf folder add root to Graph.
    # This is the starting point of the tree traversal.
    if($RootChildDirs.length -gt 0){

        $Graph.Add($SourceDir, $RootChildDirs)

    }
    # If root is a leaf folder, add the root folder to [LeafNodes_Lst] (is object-type list!)
    else{

        $Root = Get-Item -LiteralPath $SourceDir
        $LeafNodes_Lst.Add($Root)
        $LeafNodes_HashTable.Add($SourceDir,0)

        Write-Debug "GetObjects:Initialize-Graph [Root is Terminal Node]" 
    }


    ### So far we've ignored loose files at root.

    #===========================================================================
    # Get [Dir]---[Sub dir] edges
    #===========================================================================

    Write-Debug "GetObjects: [Inserting edges]" 

    foreach($Directory in $AllNodes_ObjArr){

        $DirectoryPath = $Directory.FullName
        Write-Debug "GetObjects: [Node] $DirectoryPath" 

        # Get all direct sub-folders of [$Directory]
        $ChildDirs = (Get-ChildItem -LiteralPath $DirectoryPath -Directory)

        <# 
        .NOTES
        A directory can be parent to a leaf node.
        
        Since _all_ of a directory's children are added to [ChildDirs_Lst] here,
        [ChildDirs_Lst] will contain all leaf nodes.
        This has to be considered later.
        #>

        # If $[Directory] is not a leaf node...
        if($ChildDirs.Count -gt 0) {
            
            # ...add all Folder-Subfolder edges of the folder to the graph.
            $null = $Graph.Add($DirectoryPath,$ChildDirs)
            
            Write-Debug "GetObjects: [Node] Is internal node" 
        }
        else{
            # This will _not_ exclude leaf nodes from [Graph].
            # This is merely done to check if a node is a leaf node in <Get-Objects>
            $null = $LeafNodes_Lst.Add($Directory)

            $LeafNodes_HashTable.Add($Directory.FullName,0)

            Write-Debug "GetObjects: [Node] Is leaf node" 
        }
    }

    # Add [LeafNodes] to [Graph] for convenience.
    $null = $Graph.Add("LeafNodes",$LeafNodes_Lst)
    $null = $Graph.Add("LeafNodes_HashTable", $LeafNodes_HashTable)

    return $Graph
}



function Get-Objects{
    <# 
    .SYNOPSIS
    Does a BFS on a Graph and returns a hashtable.

    .DESCRIPTION
    Does a BFS on a Graph and returns a nested hashtable
    called [DiscoveredObjects].

    [DiscoveredObjects]
    Hashtable layout:

        DiscoveredObjects = @{
            SrcDirTree = @{};
            SupportedObjects = [List[Object]];
            ExcludedObjects = @{}
            UnsupportedFolderCounter = [INT];
            UnsupportedFileCounter = [INT]
        }

    Files are analyzed for their extensions and
    folders for their content type to match certain criteria.
    
    Objects that match are stored in a list, which is
    stored as [SupportedObjects] in [DiscoveredObjects].

    The properties of the objects that don't match are stored
    in a hashtable, which is stored as [ExcludedObjects] in [DiscoveredObjects].

    .OUTPUTS
    [DiscoveredObjects] Hashtable
    #>
    Param(
        [Parameter(Mandatory)]
        [string]$SourceDir,

        [Parameter(Mandatory)]
        [hashtable]$PathsProgram,

        [Parameter(Mandatory)]
        [string]$LibraryName
    )

    $UnsupportedFileCounter = 0
    $UnsupportedFolderCounter = 0

    # $Counter[0] = Counting objects with wrong (unsupported) extensions
    # $Counter[1] = Counting "bad folders"
    $Counter = @(0,0)

    # CD-Folder
    # A folder that only contains elements of AllowedComicData is called a ComicData/CD-Folder
    # We include .txt in case a CD-folder contains a text-file with additional information about the comic.
    $AllowedComicData = [System.Collections.Generic.HashSet[String]] @('.jpg', '.jpeg', '.png', '.txt')
    $script:AllowedArchives = [System.Collections.Generic.HashSet[String]] @('.zip', '.rar', '.cbz', '.cbr')

    $Graph = Initialize-Graph -SourceDir $SourceDir
    
    # As described above 
    $DiscoveredObjects = @{}

    # List of objects to be processed by HSort.ps1
    $SupportedObjects = [List[object]]::new()
    $ExcludedObjects = @{}

    $Queue = [Queue[string]]::new()

    # Pruned SourceDirectory-Tree
    # Pruned means that we exclude CD-folders.
    $SrcDirTree = [ordered]@{}


    # Import [VisitedObjects].
    # Objects in [VisitedObjects] will not be added to [SupportedObjects] again.
    try{
        $VisitedObjects = Import-Clixml -Path "$($PathsProgram.LibFiles)\$LibraryName\VisitedObjects $LibraryName.xml"
        Write-Information -MessageData "[INFORMATION] VisitedObjects.xml imported" -InformationAction Continue
    }
    catch{
        $VisitedObjects = @{}
        Write-Information -MessageData "[INFORMATION] VisitedObjects.xml not found. Continuing..." -InformationAction Continue   
    }
    
    #===========================================================================
    # Begin BFS
    #===========================================================================
    
    Write-Debug "GetObjects: [Start BFS]"
    
    $LeafNodes_HashTable = $Graph.LeafNodes_HashTable
    
    # We start at root (the SourceDir from Settings.txt)
    $null = $Queue.Enqueue($SourceDir)

    Write-Debug "GetObjects: [SourceDir] $SourceDir"

    # Check if [SrcDir] contains folders which is here equivalent to  (-not $LeafNodes_HashTable.Contains($SourceDir))

    if (-not $LeafNodes_HashTable.ContainsKey($SourceDir)){
        
        # Every directory is visited, EXCEPT leaf nodes.
        while($Queue.Count -gt 0){    
            
            $NodePath = $Queue.Dequeue()

            $RP = Get-RelativePath -Path $NodePath -StartDir $SrcDirName
            $SrcDirTree.Add($RP, ([List[string]]::new()))

            Write-Debug "GetObjects: [Current node] $NodePath"
    
            #===========================================================================
            # Get loose files in current node
            #===========================================================================

            Get-Files -SupportedObjects $SupportedObjects -ExcludedObjects $ExcludedObjects -VisitedObjects $VisitedObjects -Counter $Counter
    
            #===========================================================================
            # Enqueue directories in NodePath (its sub-directories)
            # --- We omit leaf nodes ---
            #===========================================================================
            
            # Iterate over the current nodes children ($Graph.$NodePath is a List of direct neighbours of NodePath).
            Write-Debug "GetObjects: [Iterating over children]"
            foreach ($Child in ($Graph.$NodePath)){

                $ChildNodePath = $Child.FullName

                if ( -not $LeafNodes_HashTable.ContainsKey($ChildNodePath) ) {
                    
                    $Queue.Enqueue($ChildNodePath)
                    
                    # We do _not_ add leaf nodes to [SrcDirTree] here, since we want to prune [SrcDirTree] to suit our requirements.
                    # Some leaf nodes will be added back - See: Adding Leaf-Nodes -> Non-CD nodes
                    $SrcDirTree.$RP.Add($Child.BaseName)
                    
                    Write-Debug "GetObjects: [Child node] Is internal node"
                }
                else {
                    Write-Debug "GetObjects: [Child node] Is leaf node"
                }
                
                Write-Debug "GetObjects: [Child node] [ $NodePath ] --- [ $(($Graph.$NodePath).BaseName) ]"
            }
        }
    }
    else{

        Write-Debug "GetObjects:Get-Objects [Source is Terminal Node]"

        $RP = Get-RelativePath -Path $SrcDirName -StartDir $SrcDirName
        $SrcDirTree.Add($RP, ([List[string]]::new()))
    }

    Write-Debug "GetObjects: [Processing Leaf Nodes]"

    #===========================================================================
    # Adding Leaf-Nodes
    #===========================================================================
    
    <#
        .NOTES
        IMPORTANT
        When encountering a leaf-folder with mixed content,
        we add files that have an allowed extension to $SupportedObjects;
        but we do NOT add the folder itself!
    #>

    $CurrentExtensions =  [HashSet[string]]::new()
    $TemporarilyExcluded = [List[object]]::new()

    # NOTE
    # [$Graph.LeafNodes] will contain [SrcDir], if [SrcDir] is Leaf -> See:  <Initialize-Graph>
    foreach($LeafNode in $Graph.LeafNodes){

        $LeafPath = $LeafNode.FullName

        if (  ! $VisitedObjects.ContainsKey($LeafPath)  ) { # < A leaf node is added to visitedObjects IFF it only contains image-files

            # Get all files in [LeafNode]
            $Files = (Get-ChildItem -LiteralPath $LeafPath)
            
            if($Files.Count -gt 0){

                foreach ($File in $Files) {
            
                    $FileExtension = $File.Extension
                    $null = $CurrentExtensions.Add($FileExtension)
                    
                    # Check if a file in [LeafNode] was already visited
                    if (! $VisitedObjects.ContainsKey($File.FullName)) {

                        if ($AllowedArchives.Contains($FileExtension)) {
                            $null = $VisitedObjects.Add("$($File.FullName)", "0")
                            $null = $SupportedObjects.Add($File)
                        }
                        else{
                            # We do not add $File to [VisitedObjects] yet, since we don't want to add image-files of CD-folders.
                            $TemporarilyExcluded.Add($File)
                        }
                    }
                }

                # Add [LeafNode] itself to [SupportedObjects] IFF it only contains image-files and is not empty.
                # This only holds for CD-Folders.
                # CD-Folders will _not_ be scanned for updates once discovered.

                if ($CurrentExtensions.isSubsetOf($AllowedComicData)) { # <<< CD Node
    
                    $null = $VisitedObjects.Add($LeafPath, "0")
                    $null = $SupportedObjects.Add($LeafNode)
    
                }
                else{ # <<< Non-CD Node

                    Write-Debug "GetObjects: [Processing Non-CD Nodes] $LeafPath"

                    # === BEGIN: Add leaf nodes ===
                    # Now we add those leaf nodes to [SrcDirTree] that are Non-CD nodes.
                    #
                    $LeafParentPath = (($LeafNode.Parent).FullName)
                    $LeafParentRP = Get-RelativePath -Path ($LeafParentPath) -StartDir $SrcDirName

                    if (-not $SrcDirTree.Contains($LeafParentRP)) {

                            $SrcDirTree.Add($LeafParentRP, [List[String]]::new())
                            $SrcDirTree.$LeafParentRP.Add($LeafNode.BaseName)
                    }
                    else{

                            $SrcDirTree.$LeafParentRP.Add($LeafNode.BaseName)
                    }
                    #
                    # === END: Add leaf nodes ===

                    foreach($File in $TemporarilyExcluded){

                        $Counter[0]++

                        $FilePath = $File.FullName
                        $ParentPath = $LeafPath
                        
                        $RelativePath = Get-RelativePath -Path $FilePath -StartDir $SrcDirName
                        $RelativeParentPath = Get-RelativePath -Path ($ParentPath) -StartDir $SrcDirName
                        
                        $ExcludedObjectProperties = @{
                            Path         = $FilePath;
                            RelativePath = $RelativePath;
                            ParentPath = $ParentPath;
                            RelativeParentPath = $RelativeParentPath;
                            ObjectName   = $File.Name;
                            Reason       = "UnsupportedFile";
                            Extension    = $File.Extension
                        }
                        
                        $null = $VisitedObjects.Add($FilePath, "0")
                        $null = $ExcludedObjects.Add($FilePath, $ExcludedObjectProperties)
                                           
                    }
                }

                # A Non-CD Folder itself :
                 # Is not added to SupportedObjects
                 # Is not added to VisitedObjects
                 # Is not added to ExcludedObjects  

            }

            # Any empty folder will be ignored and:
             # Is not added to SupportedObjects
             # Is not added to VisitedObjects
             # Is not added to ExcludedObjects

        }

        $CurrentExtensions.Clear()
        $TemporarilyExcluded.Clear()
        
    }

    $UnsupportedFileCounter = $Counter[0]

    Write-Debug "GetObjects: Wrong extension array counter $($Counter[0])"

    Export-SrcDirTree -LibraryName $LibraryName -SourceDir $SourceDir -SrcDirTree $SrcDirTree
    
    #===========================================================================
    # Return objects
    #===========================================================================

    $DiscoveredObjects.Add("SrcDirTree", $SrcDirTree)
    $DiscoveredObjects.Add("SupportedObjects",$SupportedObjects)
    $DiscoveredObjects.Add("ExcludedObjects",$ExcludedObjects)

    $DiscoveredObjects.Add("UnsupportedFolderCounter", $UnsupportedFolderCounter)
    $DiscoveredObjects.Add("UnsupportedFileCounter", $UnsupportedFileCounter)

    #$Graph.Clear()

    $VisitedObjects |  Export-Clixml -Path "$($PathsProgram.LibFiles)\$LibraryName\VisitedObjects $LibraryName.xml" -Force

    return $DiscoveredObjects
}

Export-ModuleMember -Function Initialize-Graph,Get-Objects,Get-RelativePath -Variable DiscoveredObjects,ExcludedObjects,SupportedObjects,UnsupportedFileCounter,UnsupportedFolderCounter