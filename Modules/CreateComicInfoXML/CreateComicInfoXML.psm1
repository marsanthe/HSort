
function New-ComicInfoFile{
<# 
.NOTES
    16/04/2024
    Kavita 0.8 changed how collections work.

#>
    Param(
        [Parameter(Mandatory)]
        [hashtable]$ObjectMeta,

        [string]$Path
    )

    begin {
        if ($PSBoundParameters.ContainsKey('Path')) {
            $Target = $Path
        }
        else {
            $Target = $ObjectMeta.ObjectTarget
        }

        [string]$PublishingType = $ObjectMeta.PublishingType
    }

    process{

        if($PublishingType -eq "Anthology"){

            # Create and Set Formatting with XmlWriterSettings class
            $xmlObjectsettings = New-Object System.Xml.XmlWriterSettings
            $xmlObjectsettings.Indent = $true
            $xmlObjectsettings.IndentChars = "  "
            
            # Set Object path and Create document
            $XmlObjectPath = ($Target + '\ComicInfo.xml')
            $XmlObjectWriter = [System.XML.XmlWriter]::Create($XmlObjectPath, $xmlObjectsettings)
            
            # Write XML declaration and Set XSL
            $XmlObjectWriter.WriteStartDocument()
            
            # Start the Root Element and build with child nodes
            $XmlObjectWriter.WriteStartElement("ComicInfo") # <-- BaseSettings

                $XmlObjectWriter.WriteAttributeString("xmlns", "xsi", $null, "http://www.w3.org/2001/XMLSchema-instance")
                $XmlObjectWriter.WriteAttributeString("xmlns", "xsd", $null, "http://www.w3.org/2001/XMLSchema")
        
                $XmlObjectWriter.WriteElementString("Title", ($ObjectMeta.Title))
                $XmlObjectWriter.WriteElementString("Series", ($ObjectMeta.Series)) # $Title
                $XmlObjectWriter.WriteElementString("Tags", ($ObjectMeta.Tags))
                $XmlObjectWriter.WriteElementString("Format", ($ObjectMeta.Format))
                $XmlObjectWriter.WriteElementString("SeriesGroup", ($ObjectMeta.PublishingType)) # Kavita creates groups by the SeriesGroup tag Grouping by PublishingType
                $XmlObjectWriter.WriteElementString("Manga", "Yes")
          
            $XmlObjectWriter.WriteEndElement() # <-- End BaseSettings 
         
            $XmlObjectWriter.WriteEndDocument()
            $XmlObjectWriter.Flush()
            $XmlObjectWriter.Close()            
            
        }
        elseif ($PublishingType -eq "Manga") {
   
            $xmlObjectsettings = New-Object System.Xml.XmlWriterSettings

            $xmlObjectsettings.Indent = $true
            $xmlObjectsettings.IndentChars = "  "
         
            $XmlObjectPath = ($Target + '\ComicInfo.xml')
            $XmlObjectWriter = [System.XML.XmlWriter]::Create($XmlObjectPath, $xmlObjectsettings) #!
         
            $XmlObjectWriter.WriteStartDocument()
         
            $XmlObjectWriter.WriteStartElement("ComicInfo") 

                $XmlObjectWriter.WriteAttributeString("xmlns", "xsi", $null, "http://www.w3.org/2001/XMLSchema-instance")
                $XmlObjectWriter.WriteAttributeString("xmlns", "xsd", $null, "http://www.w3.org/2001/XMLSchema")
        
                $XmlObjectWriter.WriteElementString("Title", ($ObjectMeta.Title))
                $XmlObjectWriter.WriteElementString("Series", ($ObjectMeta.Series)) # Series is Title
                $XmlObjectWriter.WriteElementString("Writer", ($ObjectMeta.Artist)) # Kavita only knows writer, so writer:=artist
                $XmlObjectWriter.WriteElementString("Tags", ($ObjectMeta.Tags))
                $XmlObjectWriter.WriteElementString("Format", ($ObjectMeta.Format))
                $XmlObjectWriter.WriteElementString("SeriesGroup", ($ObjectMeta.SeriesGroup)) # Kavita creates groups by the SeriesGroup tag, we want to group by artists
                $XmlObjectWriter.WriteElementString("Manga", "Yes")
         
         
            $XmlObjectWriter.WriteEndElement() 
         
            # Finally close the XML Document
            $XmlObjectWriter.WriteEndDocument()
            $XmlObjectWriter.Flush()
            $XmlObjectWriter.Close()       

        }
        elseif ($PublishingType -eq "Doujinshi") {

            $xmlObjectsettings = New-Object System.Xml.XmlWriterSettings
            $xmlObjectsettings.Indent = $true
            $xmlObjectsettings.IndentChars = "  "
         
            $XmlObjectPath = ($Target + '\ComicInfo.xml')
            $XmlObjectWriter = [System.XML.XmlWriter]::Create($XmlObjectPath, $xmlObjectsettings)
         
            $XmlObjectWriter.WriteStartDocument()
         
            $XmlObjectWriter.WriteStartElement("ComicInfo") # <-- BaseSettings

                $XmlObjectWriter.WriteAttributeString("xmlns", "xsi", $null, "http://www.w3.org/2001/XMLSchema-instance")
                $XmlObjectWriter.WriteAttributeString("xmlns", "xsd", $null, "http://www.w3.org/2001/XMLSchema")
        
                $XmlObjectWriter.WriteElementString("Title", ($ObjectMeta.Title))
                $XmlObjectWriter.WriteElementString("Series", ($ObjectMeta.Series)) # Series is Title 
                $XmlObjectWriter.WriteElementString("Writer", ($ObjectMeta.Artist)) # Kavita only knows writer, so writer:=artist
                $XmlObjectWriter.WriteElementString("Tags", ($ObjectMeta.Tags))
                $XmlObjectWriter.WriteElementString("Format", ($ObjectMeta.Format)) # Set format to One-Shot (for now...)
                $XmlObjectWriter.WriteElementString("SeriesGroup", ($ObjectMeta.Convention)) # Group by ConventionName
                $XmlObjectWriter.WriteElementString("Manga", "Yes")
          
            $XmlObjectWriter.WriteEndElement()
         
            # Finally close the XML Document
            $XmlObjectWriter.WriteEndDocument()
            $XmlObjectWriter.Flush()
            $XmlObjectWriter.Close()
        }
    }   
}

Export-ModuleMember -Function New-ComicInfoFile