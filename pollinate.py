#!/usr/bin/env python3

import shutil
import sys
import tempfile
import zipfile
import argparse

drawing = """<w:drawing mc:Ignorable="w14 wp14">
	<wp:inline distT="0" distB="0" distL="0" distR="0">
		<wp:extent cx="1" cy="1"/>
		<wp:effectExtent l="0" t="0" r="0" b="0"/>
		<wp:docPr id="4" name="pengwings"/>
		<wp:cNvGraphicFramePr>
			<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
		</wp:cNvGraphicFramePr>
		<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
			<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
				<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
					<pic:nvPicPr>
						<pic:cNvPr id="0" name="cumberbatch"/>
						<pic:cNvPicPr/>
					</pic:nvPicPr>
					<pic:blipFill>
						<a:blip r:link="rId1337"/>
						<a:stretch>
							<a:fillRect/>
						</a:stretch>
					</pic:blipFill>
					<pic:spPr>
						<a:xfrm>
							<a:off x="0" y="0"/>
							<a:ext cx="1" cy="1"/>
						</a:xfrm>
						<a:prstGeom prst="rect">
							<a:avLst/>
						</a:prstGeom>
					</pic:spPr>
				</pic:pic>
			</a:graphicData>
		</a:graphic>
	</wp:inline>
</w:drawing>
</w:body>"""

def main():
    parser = argparse.ArgumentParser(description="Add a pingback via HTTP oAdd a pingback via HTTP to a Microsoft Office document")
    parser.add_argument("-u", "--url", type=str, help="Set a URL for the beacon.", required=True)
    parser.add_argument("-f", "--file", help="File to modified", required=True)
    parser.add_argument("-o", "--output", help="Modified output file", required=True)
    args = parser.parse_args()

    inputFile = args.file
    pingbackUrl = args.url
    outputFile = args.output

    tempdir = tempfile.mkdtemp()
    tempdir2 = tempfile.mkdtemp()

    zf = zipfile.ZipFile(inputFile, 'r')

    try:
        zf.extractall(tempdir)
    except KeyError:
        print(f"ERROR: {KeyError}")
    except:
        print("Aw snap, something else went wrong!")
    finally:
        zf.close()

    with open(tempdir+"/word/document.xml","r") as f:
        docData = f.read()

    with open(tempdir+"/word/_rels/document.xml.rels","r") as f:
        relsData = f.read()

    docData = docData.replace('</w:body>',drawing)
    relsData = relsData.replace('</Relationships>','<Relationship Id="rId1337" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="'+pingbackUrl+'" TargetMode="External"/></Relationships>')

    with open(tempdir+"/word/document.xml","w") as f:
        f.write(docData)

    with open(tempdir+"/word/_rels/document.xml.rels","w") as f:
        f.write(relsData)

    try:
        shutil.make_archive(tempdir2+"/out", "zip", tempdir)
        shutil.move(tempdir2+"/out.zip", outputFile)
        print(f"Pingback {outputFile} file created!")
        shutil.rmtree(tempdir)
        shutil.rmtree(tempdir2)
    except:
        print("Aw snap, something went wrong!")
    
    sys.exit(0)

if __name__ == "__main__":
    main()