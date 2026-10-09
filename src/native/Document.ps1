# Native Open XML builder. No Python, downloads or Office are used here.
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
function Escape-Xml($Text) { [Security.SecurityElement]::Escape([string]$Text) }
function New-Run($Text, $Font='宋体', $Size=24, $Bold=$false) {
    $b = if ($Bold) { '<w:b/>' } else { '<w:b w:val="0"/>' }
    '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="'+$Font+'"/>'+$b+'<w:sz w:val="'+$Size+'"/><w:color w:val="000000"/></w:rPr><w:t xml:space="preserve">'+(Escape-Xml $Text)+'</w:t></w:r>'
}
function New-Paragraph($Text, $Kind='body', $Bookmark='', $Extra='') {
    $font='宋体'; $size=24; $bold=$false; $indent=0; $align='left'; $style=''; $keep=''
    switch ($Kind) {
        body { $indent=200; $style='<w:pStyle w:val="ReportNativeBody"/>' }
        h1 { $font='黑体'; $size=32; $bold=$true; $style='<w:pStyle w:val="ReportNativeH1"/>'; $keep='<w:keepNext/>' }
        h2 { $font='黑体'; $size=30; $bold=$true; $style='<w:pStyle w:val="ReportNativeH2"/>'; $keep='<w:keepNext/>' }
        toc1 { $bold=$true }
        toc2 { $size=21 }
        center { $align='center' }
        title { $font='黑体'; $size=32; $bold=$true; $align='center' }
    }
    $mark=''; $end=''
    if ($Bookmark) { $script:bookmarkId++; $mark='<w:bookmarkStart w:id="'+$script:bookmarkId+'" w:name="'+$Bookmark+'"/>'; $end='<w:bookmarkEnd w:id="'+$script:bookmarkId+'"/>' }
    '<w:p><w:pPr>'+$style+$keep+'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8107"/></w:tabs><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="'+($indent*2.4)+'" w:firstLineChars="'+$indent+'"/><w:jc w:val="'+$align+'"/></w:pPr>'+$mark+(New-Run $Text $font $size $bold)+$Extra+$end+'</w:p>'
}
function New-Field($Code) { '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> '+(Escape-Xml $Code)+' </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'+(New-Run '0')+'<w:r><w:fldChar w:fldCharType="end"/></w:r>' }
function Get-Headings($Data) {
    $list=@(); $cn=@('一','二','三','四','五','六','七','八','九'); $i=0; $globalSub=0
    foreach ($section in $Data.sections) {
        $i++; $list+=@{body="$($cn[$i-1])、$($section.title)";toc="$i. $($section.title)";bookmark="_ReportH1$i";level=1;node=$section;section=$i;sub=0}
        $j=0
        foreach ($sub in $section.subsections) { $j++; $globalSub++; $list+=@{body="$i.$j $($sub.title)";toc="$i.$j $($sub.title)";bookmark="_ReportH2${i}_$globalSub";level=2;node=$sub;section=$i;sub=$j} }
    }
    if ($Data.code_block_mode -eq 'after_body') { $i++; $list+=@{body="$($cn[$i-1])、代码块";toc="$i. 代码块";bookmark="_ReportH1$i";level=1;node=@{paragraphs=@()};section=$i;sub=0;code=$true} }
    return ,$list
}
function Read-ZipText($Zip,$Name) {
    $entry=$Zip.GetEntry($Name); if (-not $entry) { throw "Missing DOCX part: $Name" }
    $reader=[IO.StreamReader]::new($entry.Open()); try { $reader.ReadToEnd() } finally { $reader.Dispose() }
}
function Write-ZipText($Zip,$Name,$Text) {
    $old=$Zip.GetEntry($Name); if ($old) { $old.Delete() }
    $entry=$Zip.CreateEntry($Name); $writer=[IO.StreamWriter]::new($entry.Open(),[Text.UTF8Encoding]::new($false))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}
function Read-SafeXml($Text) {
    $settings=[Xml.XmlReaderSettings]::new(); $settings.DtdProcessing=[Xml.DtdProcessing]::Prohibit; $settings.XmlResolver=$null
    $reader=[Xml.XmlReader]::Create([IO.StringReader]::new($Text),$settings)
    try { $doc=[xml]::new(); $doc.XmlResolver=$null; $doc.Load($reader); return ,$doc } finally { $reader.Dispose() }
}
function Replace-ParagraphText($Paragraph,$Ns,$NewText) {
    # Replace the slot across runs while retaining each run's formatting and surrounding structure.
    $texts=@($Paragraph.SelectNodes('.//w:t',$Ns))
    if (-not $texts.Count) {throw '模板字段缺少可填写的文本。'}
    $texts[0].InnerText=$NewText
    for($i=1;$i -lt $texts.Count;$i++){$texts[$i].InnerText=''}
}
function Fill-FrontMatter($Document,$Ns,$Metadata) {
    $labels=Get-FrontMatterLabels $Metadata
    $patterns=@{
        score_title='《[^》]+》\s*实践周成绩报告单';cover_title='《[^》]+》\s*实践周总结'
        score_term='\d{3,4}[xX]?\s*[-—–－]\s*\d{3,4}[xX]?\s*学年度\s*第[一二12xX]学期'
        cover_term='\d{3,4}[xX]?\s*[-—–－]\s*\d{3,4}[xX]?\s*学年\s*第[一二12xX]学期\s*期末考试'
    }
    $names=@{score_title='_QBScoreTitle';cover_title='_QBCoverTitle';score_term='_QBScoreTerm';cover_term='_QBCoverTerm'}
    foreach($key in @('score_title','cover_title','score_term','cover_term')){
        $found=@()
        foreach($p in $Document.SelectNodes('//w:body//w:p',$Ns)){
            $text=($p.SelectNodes('.//w:t',$Ns) | ForEach-Object {$_.InnerText}) -join ''
            if($text -match $patterns[$key]){$found+=,$p}
        }
        if($found.Count -ne 1){throw "学校模板字段必须能唯一定位：$key（找到$($found.Count)处）。"}
        $p=$found[0];$old=($p.SelectNodes('.//w:t',$Ns) | ForEach-Object {$_.InnerText}) -join ''
        $replacement=[regex]::Replace($old,$patterns[$key],[Text.RegularExpressions.MatchEvaluator]{param($m) $labels[$key]})
        Replace-ParagraphText $p $Ns $replacement
        $script:bookmarkId++
        $start=$Document.CreateElement('w','bookmarkStart',$Ns.LookupNamespace('w'));[void]$start.SetAttribute('id',$Ns.LookupNamespace('w'),[string]$script:bookmarkId);[void]$start.SetAttribute('name',$Ns.LookupNamespace('w'),$names[$key])
        $end=$Document.CreateElement('w','bookmarkEnd',$Ns.LookupNamespace('w'));[void]$end.SetAttribute('id',$Ns.LookupNamespace('w'),[string]$script:bookmarkId)
        [void]$p.AppendChild($start);[void]$p.AppendChild($end)
    }
}
function Build-Document($Data,$Path,$Template) {
    $script:bookmarkId=10000
    $w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'; $r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $page='<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1797" w:bottom="1440" w:left="1797" w:header="850" w:footer="850" w:gutter="0"/>'
    $front=''; $rels='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'; $types='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>'
    if ($Template) { [IO.File]::Copy($Template,$Path,$false) }
    else { $file=[IO.File]::Create($Path); $file.Dispose() }
    $zip=[IO.Compression.ZipFile]::Open($Path,[IO.Compression.ZipArchiveMode]::Update)
    try {
        if ($Template) {
            $doc=Read-SafeXml (Read-ZipText $zip 'word/document.xml'); $ns=[Xml.XmlNamespaceManager]::new($doc.NameTable); $ns.AddNamespace('w',$w)
            if ($doc.SelectNodes('//w:sectPr',$ns).Count -gt 1) { throw 'School template has multiple sections; normalize the three-page template before native generation.' }
            if($Data.front_matter){Fill-FrontMatter $doc $ns $Data.front_matter}
            foreach ($node in $doc.SelectNodes('//w:body/*[not(self::w:sectPr)]',$ns)) { $front+=$node.OuterXml }
            # Preserve template relationships, drawings and styles. Never touch the source template.
            $rels=Read-ZipText $zip 'word/_rels/document.xml.rels'; $types=Read-ZipText $zip '[Content_Types].xml'
            $rootOpen=$doc.DocumentElement.OuterXml.Split('>')[0]+'>'
        } else {
            $rootOpen='<w:document xmlns:w="'+$w+'" xmlns:r="'+$r+'">'
            if($Data.front_matter){
                $labels=Get-FrontMatterLabels $Data.front_matter
                $front=(New-Paragraph $labels.score_term plain '_QBScoreTerm')+(New-Paragraph $labels.score_title title '_QBScoreTitle')+(New-Paragraph '成绩：________    教师：________' plain)
                $term=New-Paragraph $labels.cover_term plain '_QBCoverTerm';$term=$term.Replace('<w:pPr>','<w:pPr><w:pageBreakBefore/>')
                $front+=$term+(New-Paragraph $labels.cover_title title '_QBCoverTitle')+(New-Paragraph '请替换为学校指定封面。' plain)
            }else{
                $front=(New-Paragraph '成绩报告单样板' title)+(New-Paragraph '请替换为学校指定的成绩报告单。' plain)+(New-Paragraph '成绩：________    教师：________' plain)
                $front+='<w:p><w:pPr><w:pageBreakBefore/></w:pPr>'+(New-Run '封面样板')+'</w:p>'+(New-Paragraph $Data.report_title title)+(New-Paragraph '请替换为学校指定封面。' plain)
            }
            $front+='<w:p><w:pPr><w:pageBreakBefore/></w:pPr>'+(New-Run '撰写要求样板')+'</w:p>'+(New-Paragraph '请替换为学校指定的撰写要求。' plain)
        }
        $toc=New-Paragraph '目录' title
        $toc=$toc.Replace('<w:pPr>','<w:pPr><w:pageBreakBefore/>')
        $heads=Get-Headings $Data
        foreach ($h in $heads) { $toc+=New-Paragraph $h.toc "toc$($h.level)" '' ('<w:r><w:tab/></w:r>'+(New-Field "PAGEREF $($h.bookmark) \h")) }
        # Put section boundary on the last TOC entry; no empty spacer paragraph.
        $last=$toc.LastIndexOf('</w:pPr>'); $toc=$toc.Insert($last,'<w:sectPr>'+$page+'</w:sectPr>')
        $body=''
        foreach ($h in $heads) {
            $body+=New-Paragraph $h.body "h$($h.level)" $h.bookmark
            foreach ($p in $h.node.paragraphs) { $body+=New-Paragraph $p }
            $blocks=@()
            if ($h.code) { $blocks=@($Data.code_blocks) }
            elseif ($Data.code_block_mode -eq 'embedded') { $blocks=@($Data.code_blocks | Where-Object { $_.section_index -eq $h.section -and [int]$_.subsection_index -eq $h.sub }) }
            foreach ($block in $blocks) { if ($block.caption) { $body+=New-Paragraph $block.caption plain }; foreach ($line in ($block.content -split '\r?\n')) { if ($line.Trim()) { $body+=New-Paragraph $line plain } } }
            if ($h.node.image_placeholder) { $body+=New-Paragraph "【图片预留位置：$($h.node.image_placeholder)】" center }
        }
        $document=$rootOpen+'<w:body>'+$front+$toc+$body+'<w:sectPr><w:footerReference w:type="default" r:id="reportNativeFooter"/><w:type w:val="nextPage"/>'+$page+'<w:pgNumType w:start="1"/></w:sectPr></w:body></w:document>'
        $relDoc=Read-SafeXml $rels
        foreach ($spec in @(@('reportNativeFooter','footer','report-native-footer.xml'),@('reportNativeStyles','styles','styles.xml'))) {
            if ($spec[1] -eq 'styles' -and $zip.GetEntry('word/styles.xml')) { continue }
            $el=$relDoc.CreateElement('Relationship',$relDoc.DocumentElement.NamespaceURI); $el.SetAttribute('Id',$spec[0]); $el.SetAttribute('Type',"$r/$($spec[1])"); $el.SetAttribute('Target',$spec[2]); [void]$relDoc.DocumentElement.AppendChild($el)
        }
        $typeDoc=Read-SafeXml $types
        foreach ($spec in @(@('/word/document.xml','document.main'),@('/word/styles.xml','styles'),@('/word/report-native-footer.xml','footer'))) {
            $found=@($typeDoc.DocumentElement.ChildNodes | Where-Object { $_.GetAttribute('PartName') -eq $spec[0] })
            if (-not $found.Count) { $el=$typeDoc.CreateElement('Override',$typeDoc.DocumentElement.NamespaceURI); $el.SetAttribute('PartName',$spec[0]); $el.SetAttribute('ContentType',"application/vnd.openxmlformats-officedocument.wordprocessingml.$($spec[1])+xml"); [void]$typeDoc.DocumentElement.AppendChild($el) }
        }
        if (-not $zip.GetEntry('_rels/.rels')) { Write-ZipText $zip '_rels/.rels' ('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="'+$r+'/officeDocument" Target="word/document.xml"/></Relationships>') }
        if (-not $zip.GetEntry('word/styles.xml')) { Write-ZipText $zip 'word/styles.xml' ('<w:styles xmlns:w="'+$w+'"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style></w:styles>') }
        $styles=Read-SafeXml (Read-ZipText $zip 'word/styles.xml')
        foreach ($spec in @(@('ReportNativeBody','Report Native Body',''),@('ReportNativeH1','Report Native H1','<w:outlineLvl w:val="0"/>'),@('ReportNativeH2','Report Native H2','<w:outlineLvl w:val="1"/>'))) {
            $fragment=$styles.CreateDocumentFragment()
            $fragment.InnerXml='<w:style xmlns:w="'+$w+'" w:type="paragraph" w:styleId="'+$spec[0]+'"><w:name w:val="'+$spec[1]+'"/><w:pPr><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/><w:jc w:val="left"/>'+$spec[2]+'</w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="24"/><w:b w:val="0"/></w:rPr></w:style>'
            [void]$styles.DocumentElement.AppendChild($fragment)
        }
        Write-ZipText $zip 'word/styles.xml' $styles.OuterXml
        Write-ZipText $zip 'word/document.xml' $document
        Write-ZipText $zip 'word/_rels/document.xml.rels' $relDoc.OuterXml
        Write-ZipText $zip '[Content_Types].xml' $typeDoc.OuterXml
        Write-ZipText $zip 'word/report-native-footer.xml' ('<w:ftr xmlns:w="'+$w+'">'+(New-Paragraph '' center '' (New-Field 'PAGE'))+'</w:ftr>')
    } finally { $zip.Dispose() }
}
function Effective-Attribute($Paragraph,$Run,$Styles,$Ns,$Kind,$Element,$Attribute,$Fallback) {
    $uri=$Ns.LookupNamespace('w'); $sources=@()
    if ($Run) { $sources+=,$Run.SelectSingleNode("w:rPr/w:$Element",$Ns) }
    if ($Kind -eq 'pPr') { $sources+=,$Paragraph.SelectSingleNode("w:pPr/w:$Element",$Ns) }
    $styleId=$Paragraph.SelectSingleNode('w:pPr/w:pStyle',$Ns)
    $style=if($styleId){$Styles.SelectSingleNode("//w:style[@w:styleId='$($styleId.GetAttribute('val',$uri))']",$Ns)}else{$Styles.SelectSingleNode('//w:style[@w:type="paragraph" and @w:default="1"]',$Ns)}
    $seen=@()
    while ($style) {
        $id=$style.GetAttribute('styleId',$uri); if ($seen -contains $id) {throw 'Cyclic style inheritance.'}; $seen+=$id
        $sources+=,$style.SelectSingleNode("w:$Kind/w:$Element",$Ns)
        $based=$style.SelectSingleNode('w:basedOn',$Ns)
        $style=if($based){$Styles.SelectSingleNode("//w:style[@w:styleId='$($based.GetAttribute('val',$uri))']",$Ns)}else{$null}
    }
    $sources+=,$Styles.SelectSingleNode("//w:docDefaults/w:${Kind}Default/w:$Kind/w:$Element",$Ns)
    foreach ($node in $sources) {
        if ($node -and $node.HasAttribute($Attribute,$uri)) {return $node.GetAttribute($Attribute,$uri)}
        if ($node -and $Element -eq 'b') {return '1'}
    }
    return $Fallback
}
function Test-SavedDocument($Path,$Map,$Data) {
    $zip=[IO.Compression.ZipFile]::OpenRead($Path)
    try { $doc=Read-SafeXml (Read-ZipText $zip 'word/document.xml'); $styles=Read-SafeXml (Read-ZipText $zip 'word/styles.xml') } finally { $zip.Dispose() }
    $ns=[Xml.XmlNamespaceManager]::new($doc.NameTable); $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $heads=Get-Headings $Data
    if($Data.front_matter){
        $labels=Get-FrontMatterLabels $Data.front_matter
        $names=@{score_title='_QBScoreTitle';cover_title='_QBCoverTitle';score_term='_QBScoreTerm';cover_term='_QBCoverTerm'}
        foreach($key in $names.Keys){
            $marks=$doc.SelectNodes("//w:bookmarkStart[@w:name='$($names[$key])']",$ns)
            if($marks.Count -ne 1){throw "前置页字段书签缺失或重复：$key"}
            $text=($marks[0].ParentNode.SelectNodes('.//w:t',$ns) | ForEach-Object {$_.InnerText}) -join ''
            if($text.Trim() -ne $labels[$key]){throw "前置页题目或学年学期与确认不同：$key"}
        }
        if(-not $Map.front_matter_pages -or $Map.front_matter_pages.Count -ne 4){throw '成绩单/封面四个字段的实际页码映射缺失。'}
        foreach($name in $names.Values){$expected=if($name -like '_QBScore*'){1}else{2};if($Map.front_matter_pages[$name] -ne $expected){throw '成绩单/封面字段所在物理页不正确。'}}
    }
    if (@($Map.page_map).Count -ne $heads.Count -or $Map.status -ne 'ok') { throw 'Incomplete live pagination map.' }
    foreach ($h in $heads) {
        $marks=$doc.SelectNodes("//w:bookmarkStart[@w:name='$($h.bookmark)']",$ns)
        if ($marks.Count -ne 1) { throw "Heading bookmark is not unique: $($h.bookmark)" }
        $title=($marks[0].ParentNode.SelectNodes('.//w:t',$ns) | ForEach-Object {$_.InnerText}) -join ''
        if ($title -ne $h.body) { throw "Saved body heading differs: $title" }
        $hp=$marks[0].ParentNode
        $spacing=$hp.SelectSingleNode('w:pPr/w:spacing',$ns)
        if ((Effective-Attribute $hp $null $styles $ns pPr spacing line '') -ne '360' -or (Effective-Attribute $hp $null $styles $ns pPr spacing before '0') -ne '0' -or (Effective-Attribute $hp $null $styles $ns pPr spacing after '0') -ne '0') { throw 'Heading spacing changed.' }
        foreach ($run in $hp.SelectNodes('w:r[w:t]',$ns)) {
            $font=$run.SelectSingleNode('w:rPr/w:rFonts',$ns); $size=$run.SelectSingleNode('w:rPr/w:sz',$ns)
            $expectedSize=if($h.level -eq 1){'32'}else{'30'}
            if ((Effective-Attribute $hp $run $styles $ns rPr rFonts eastAsia '') -ne '黑体' -or (Effective-Attribute $hp $run $styles $ns rPr rFonts ascii '') -ne 'Times New Roman' -or (Effective-Attribute $hp $run $styles $ns rPr sz val '') -ne $expectedSize) { throw 'Heading font changed.' }
        }
        $rows=@($Map.page_map | Where-Object {$_.bookmark -eq $h.bookmark})
        if ($rows.Count -ne 1 -or $rows[0].toc_page -ne $rows[0].body_footer_page) { throw 'TOC/footer mismatch.' }
        $fields=$doc.SelectNodes('//w:instrText',$ns) | Where-Object {$_.InnerText -match ('PAGEREF\s+'+[regex]::Escape($h.bookmark)+'(?:\s|$)')}
        if (@($fields).Count -ne 1) { throw 'Saved PAGEREF missing or duplicated.' }
        $p=@($fields)[0].ParentNode.ParentNode
        $text=($p.SelectNodes('.//w:t',$ns) | ForEach-Object {$_.InnerText}) -join ''
        if ($text -ne ($h.toc+[string]$rows[0].toc_page)) { throw "Saved TOC title/page differs: $text" }
        foreach ($run in $p.SelectNodes('w:r[w:t]',$ns) | Select-Object -First 1) {
            $tocSize=if($h.level -eq 1){'24'}else{'21'}
            $tocBold=Effective-Attribute $p $run $styles $ns rPr b val '0'
            if ((Effective-Attribute $p $run $styles $ns rPr rFonts eastAsia '') -ne '宋体' -or (Effective-Attribute $p $run $styles $ns rPr sz val '') -ne $tocSize -or ($h.level -eq 1 -and $tocBold -notin @('1','true','on')) -or ($h.level -eq 2 -and $tocBold -notin @('0','false','off'))) { throw 'TOC font/bold changed.' }
        }
    }
    if ($Map.toc_start_physical_page -ne 4 -or $Map.body_start_physical_page -ne ($Map.toc_end_physical_page+1)) { throw 'Front matter/TOC/body pagination invalid.' }
    $inBody=$false
    foreach ($p in $doc.SelectNodes('/w:document/w:body/w:p',$ns)) {
        if ($p.SelectSingleNode("w:bookmarkStart[@w:name='_ReportH11']",$ns)) { $inBody=$true }
        if (-not $inBody -or $p.SelectSingleNode('w:bookmarkStart[starts-with(@w:name,"_ReportH")]',$ns)) { continue }
        $text=($p.SelectNodes('.//w:t',$ns) | ForEach-Object {$_.InnerText}) -join ''
        if (-not $text.Trim()) { throw 'Empty paragraph in body.' }
        $spacing=$p.SelectSingleNode('w:pPr/w:spacing',$ns)
        if ((Effective-Attribute $p $null $styles $ns pPr spacing line '') -ne '360') { throw 'Body spacing changed.' }
        foreach ($run in $p.SelectNodes('w:r[w:t]',$ns)) {
            $font=$run.SelectSingleNode('w:rPr/w:rFonts',$ns); $size=$run.SelectSingleNode('w:rPr/w:sz',$ns); $bold=$run.SelectSingleNode('w:rPr/w:b',$ns)
            if ((Effective-Attribute $p $run $styles $ns rPr rFonts eastAsia '') -ne '宋体' -or (Effective-Attribute $p $run $styles $ns rPr rFonts ascii '') -ne 'Times New Roman' -or (Effective-Attribute $p $run $styles $ns rPr sz val '') -ne '24' -or (Effective-Attribute $p $run $styles $ns rPr b val '0') -notin @('0','false','off')) { throw 'Body font/bold changed.' }
        }
    }
}
