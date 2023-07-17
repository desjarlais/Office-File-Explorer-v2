using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Windows.Forms;

namespace Office_File_Explorer.Helpers
{
    class OfficeDocument
    {
		private Package _package;
		private List<OfficePart> _xmlParts;
		private string _fileName;

		public OfficeDocument(string fileName)
        {
			_fileName = fileName;
			_package = Package.Open(fileName, FileMode.Open, FileAccess.Read);
			
			if (_package is null) return;

			_xmlParts = new List<OfficePart>();



			// clp 
			foreach (PackageRelationship pr in _package.GetRelationshipsByType(Strings.schemaClpRelationship))
			{
				_xmlParts.Add(new OfficePart(_package.GetPart(pr.TargetUri), XMLParts.LabelInfo, pr.Id));
			}

			// office 2010 custom ui parts
			foreach (PackageRelationship relationship in _package.GetRelationshipsByType(Strings.CustomUI14PartRelType))
			{
				Uri customUIUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
				if (_package.PartExists(customUIUri))
				{
					_xmlParts.Add(new OfficePart(_package.GetPart(customUIUri), XMLParts.RibbonX14, relationship.Id));
				}
				break;
			}

			// office 2007 custom ui parts
			foreach (PackageRelationship relationship in _package.GetRelationshipsByType(Strings.CustomUIPartRelType))
			{
				Uri customUIUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
				if (_package.PartExists(customUIUri))
				{
					_xmlParts.Add(new OfficePart(_package.GetPart(customUIUri), XMLParts.RibbonX12, relationship.Id));
				}
				break;
			}

			// qat parts
			foreach (PackageRelationship relationship in _package.GetRelationshipsByType(Strings.QATPartRelType))
			{
				Uri qatUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
				if (_package.PartExists(qatUri))
				{
					_xmlParts.Add(new OfficePart(_package.GetPart(qatUri), XMLParts.QAT12, relationship.Id));
				}
				break;
			}
		}

		/// <summary>
		/// need to clean up the refs
		/// </summary>
		public void ClosePackage()
        {
			_package.Close();
        }

		#region Basic Accessors
		public List<OfficePart> Parts
		{
			get
			{
				return _xmlParts;
			}
		}

		public string Name
		{
			get
			{
				return _fileName;
			}
		}

		public bool HasCustomUI
		{
			get
			{
				if (_xmlParts is null || _xmlParts.Count == 0) return false;

				for (int i = 0; i < _xmlParts.Count; i++)
				{
					if (_xmlParts[i].PartType == XMLParts.RibbonX12 || _xmlParts[i].PartType == XMLParts.RibbonX14)
					{
						return true;
					}
				}

				return false;
			}
		}
		#endregion

		public OfficePart CreateCustomPart(XMLParts partType)
		{
			string relativePath;
			string relType;

			switch (partType)
			{
				case XMLParts.fontTable:
					relativePath = "/fontTable.xml";
					relType = Strings.schemaMsft2006 + "fontTable";
					break;
                case XMLParts.Styles:
                    relativePath = "/styles.xml";
                    relType = Strings.schemaMsft2006 + "styles.xml";
                    break;
                case XMLParts.Theme:
                    relativePath = "/theme/theme.xml";
                    relType = Strings.schemaMsft2006 + "/theme/theme1.xml";
                    break;
                case XMLParts.webSettings:
                    relativePath = "/webSettings.xml";
                    relType = Strings.schemaMsft2006 + "webSettings.xml";
                    break;
                case XMLParts.Settings:
                    relativePath = "/settings.xml";
                    relType = Strings.schemaMsft2006 + "settings.xml";
                    break;
				case XMLParts.CoreProps:
                    relativePath = "/docProps/core.xml";
                    relType = Strings.schemaMsft2006 + "metadata/core-properties/docProps/core.xml";
                    break;
				case XMLParts.AppProps:
                    relativePath = "/docProps/app.xml";
                    relType = Strings.schemaMsft2006 + "extended-properties/docProps/app.xml";
                    break;
				case XMLParts.Document:
                    relativePath = "/word/document.xml";
                    relType = Strings.schemaMsft2006 + "officeDocument/word/document.xml";
                    break;
                case XMLParts.LabelInfo:
                    relativePath = "/docMetadata/LabelInfo.xml";
                    relType = Strings.schemaClpRelationship;
                    break;
                case XMLParts.RibbonX12:
					relativePath = "/customUI/customUI.xml";
					relType = Strings.CustomUIPartRelType;
					break;
				case XMLParts.RibbonX14:
					relativePath = "/customUI/customUI14.xml";
					relType = Strings.CustomUI14PartRelType;
					break;
				case XMLParts.QAT12:
					relativePath = "/customUI/qat.xml";
					relType = Strings.QATPartRelType;
					break;
				default:
					FileUtilities.WriteToLog(Strings.fLogFilePath, "CreateCustomPart Error - Unknown type");
					return null;
			}

			Uri customUIUri = new Uri(relativePath, UriKind.Relative);
			PackageRelationship relationship = _package.CreateRelationship(customUIUri, TargetMode.Internal, relType);

            OfficePart part;
            if (!_package.PartExists(customUIUri))
			{
				part = new OfficePart(_package.CreatePart(customUIUri, "application/xml"), partType, relationship.Id);
			}
			else
			{
				part = new OfficePart(_package.GetPart(customUIUri), partType, relationship.Id);
			}
			FileUtilities.WriteToLog(Strings.fLogFilePath, "Fail to create custom part.");

			_xmlParts.Add(part);

			return part;
		}

		public OfficePart RetrieveCustomPart(XMLParts partType)
		{
			if (_xmlParts is null || _xmlParts.Count == 0) return null;

			for (int i = 0; i < _xmlParts.Count; i++)
			{
				if (_xmlParts[i].PartType == partType)
				{
					return _xmlParts[i];
				}
			}
			return null;
		}
	}

	class OfficePart
	{
		XMLParts _partType;
		PackagePart _part;
		string _id;
		string _name;

		public OfficePart(PackagePart part, XMLParts partType, string relationshipId)
		{
			_part = part;
			_partType = partType;
			_id = relationshipId;
			_name = Path.GetFileName(_part.Uri.ToString());
		}

		public PackagePart Part
		{
			get
			{
				return _part;
			}
		}

		public XMLParts PartType
		{
			get
			{
				return _partType;
			}
		}

		public string Name
		{
			get
			{
				return _name;
			}
		}

		public string ReadContent()
		{
			TextReader rd = new StreamReader(_part.GetStream(FileMode.Open, FileAccess.Read));
			if (rd is null) return null;

			string text = rd.ReadToEnd();
			rd.Close();
			return text;
		}

		public void Save(string text)
		{
			if (text is null) return;

			TextWriter tw = new StreamWriter(_part.GetStream(FileMode.Create, FileAccess.Write));

			if (tw is null) return;

			tw.Write(text);
			tw.Flush();
			tw.Close();
		}

		public List<TreeNode> GetImages(ImageList imageList, ContextMenuStrip ctxMenuStrip)
		{
			if (imageList is null)
			{
				throw new ArgumentNullException("imageList");
			}

			List<TreeNode> imageCollection = new List<TreeNode>();

			foreach (PackageRelationship relationship in _part.GetRelationshipsByType(Strings.ImagePartRelType))
			{
				Uri customImageUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
				if (!_part.Package.PartExists(customImageUri)) continue;

				PackagePart imagePart = _part.Package.GetPart(customImageUri);

				Stream imageStream = imagePart.GetStream(FileMode.Open, FileAccess.Read);
				System.Drawing.Image image = System.Drawing.Image.FromStream(imageStream);

				TreeNode imageNode = new TreeNode(relationship.Id);
				imageNode.ImageKey = "_" + relationship.Id;
				imageNode.SelectedImageKey = imageNode.ImageKey;

				if (ctxMenuStrip != null)
				{
					imageNode.ContextMenuStrip = ctxMenuStrip;
				}
				imageNode.Tag = _partType;

				imageCollection.Add(imageNode);
				imageList.Images.Add(imageNode.ImageKey, image);
				imageStream.Close();
			}

			return imageCollection;
		}

		public string AddImage(string fileName, string id)
		{
			if (_partType != XMLParts.RibbonX12 && _partType != XMLParts.RibbonX14)
			{
				return null;
			}

			if (fileName is null) throw new ArgumentNullException("fileName");
			if (fileName.Length == 0) return null;

			if (id is null) throw new ArgumentNullException("id");
			if (id.Length == 0) throw new ArgumentException(Strings.idsNonEmptyId);

			if (_part.RelationshipExists(id))
			{
				id = "rId";
			}
			return AddImageHelper(fileName, id);
		}

		private string AddImageHelper(string fileName, string id)
		{
			if (fileName is null) throw new ArgumentNullException("fileName");

			FileUtilities.WriteToLog(Strings.fLogFilePath, fileName + "does not exist.");
			if (!File.Exists(fileName)) return null;

			BinaryReader br = new BinaryReader(File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));
			FileUtilities.WriteToLog(Strings.fLogFilePath, "Fail to create a BinaryReader from file.");
			if (br is null) return null;

			Uri imageUri = new Uri("images/" + Path.GetFileName(fileName), UriKind.Relative);
			int fileIndex = 0;
			while (true)
			{
				if (_part.Package.PartExists(PackUriHelper.ResolvePartUri(_part.Uri, imageUri)))
				{
					FileUtilities.WriteToLog(Strings.fLogFilePath, imageUri.ToString() + " already exists.");
					imageUri = new Uri(
						"images/" +
						Path.GetFileNameWithoutExtension(fileName) +
						(fileIndex++).ToString() +
						Path.GetExtension(fileName),
						UriKind.Relative);
					continue;
				}
				break;
			}

			if (id != null)
			{
				int idIndex = 0;
				string testId = id;
				while (true)
				{
					if (_part.RelationshipExists(testId))
					{
						FileUtilities.WriteToLog(Strings.fLogFilePath, testId + " already exists.");
						testId = id + (idIndex++);
						continue;
					}
					id = testId;
					break;
				}
			}

			PackageRelationship imageRel = _part.CreateRelationship(imageUri, TargetMode.Internal, Strings.ImagePartRelType, id);

			FileUtilities.WriteToLog(Strings.fLogFilePath, "Fail to create image relationship.");
			if (imageRel is null) return null;

			PackagePart imagePart = _part.Package.CreatePart(PackUriHelper.ResolvePartUri(imageRel.SourceUri, imageRel.TargetUri),
				MapImageContentType(Path.GetExtension(fileName)));

			FileUtilities.WriteToLog(Strings.fLogFilePath, "Fail to create image part.");
			if (imagePart is null) return null;

			BinaryWriter bw = new BinaryWriter(imagePart.GetStream(FileMode.Create, FileAccess.Write));
			FileUtilities.WriteToLog(Strings.fLogFilePath, "Fail to create a BinaryWriter to write to part.");
			if (bw is null) return null;

			byte[] buffer = new byte[1024];
			int byteCount;
			while ((byteCount = br.Read(buffer, 0, buffer.Length)) > 0)
			{
				bw.Write(buffer, 0, byteCount);
			}

			bw.Flush();
			bw.Close();
			br.Close();

			return imageRel.Id;
		}

		public void RemoveImage(string id)
		{
			if (id is null) throw new ArgumentNullException("id");
			if (id.Length == 0) return;

			if (!_part.RelationshipExists(id)) return;

			PackageRelationship imageRel = _part.GetRelationship(id);

			Uri imageUri = PackUriHelper.ResolvePartUri(imageRel.SourceUri, imageRel.TargetUri);
			if (_part.Package.PartExists(imageUri))
			{
				_part.Package.DeletePart(imageUri);
			}

			_part.DeleteRelationship(id);
		}

		public void Remove()
		{
			// Remove all image parts first
			foreach (PackageRelationship relationship in _part.GetRelationships())
			{
				Uri relUri = PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri);
				if (_part.Package.PartExists(relUri))
				{
					_part.Package.DeletePart(relUri);
				}
			}

			_part.Package.DeleteRelationship(_id);
			_part.Package.DeletePart(_part.Uri);

			_part = null;
			_id = null;
		}

		public void ChangeImageId(string source, string target)
		{
			if (source is null) throw new ArgumentNullException("source");
			if (target is null) throw new ArgumentNullException("target");
			if (target.Length == 0) throw new ArgumentException(Strings.idsNonEmptyId);

			if (source == target)
			{
				return;
			}

			if (!_part.RelationshipExists(source)) return;
			if (_part.RelationshipExists(target))
			{
				throw new Exception(Strings.idsDuplicateId.Replace("|1", target));
			}

			PackageRelationship imageRel = _part.GetRelationship(source);

			_part.CreateRelationship(imageRel.TargetUri, imageRel.TargetMode, imageRel.RelationshipType, target);
			_part.DeleteRelationship(source);
		}

		private static string MapImageContentType(string extension)
		{
			if (extension is null) throw new ArgumentNullException("extension");
			if (extension.Length == 0) throw new ArgumentException("Extension cannot be empty.");

			string extLowerCase = extension.ToLower();

			switch (extLowerCase)
			{
				case "jpg":
					return "image/jpeg";
				default:
					return "image/" + extLowerCase;
			}
		}
	}

	public enum XMLParts
	{
		webSettings,
		Settings,
		Styles,
		fontTable,
		Theme,
		AppProps,
		CoreProps,
		Document,
		LabelInfo,
		QAT12,
		RibbonX12,
		RibbonX14,
		LastEntry //Always Last
	}
}
