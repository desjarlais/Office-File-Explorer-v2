﻿using System.Text;
using System.Xml;

namespace Office_File_Explorer.Helpers
{
    class CallbackBuilder
    {
		/// <summary>
		/// Generates callbacks for given custom UI Xml.
		/// </summary>
		/// <param name="customUIXml">The custom UI Xml to generate callbacks for.</param>
		/// <returns>List of callbacks.</returns>
		public static StringBuilder GenerateCallback(XmlDocument customUIXml)
		{
			StringBuilder result = new StringBuilder();
			result.Append(Strings.rtfString);
			if (attributeList is null)
			{
				attributeList = new System.Collections.Hashtable();
			}
			attributeList.Clear();
			GenerateCallback(customUIXml.DocumentElement, result);
			return ColorizingCallbacks(result);
		}

		private static void GenerateCallback(XmlNode node, StringBuilder result)
		{
			if (node.Attributes is not null)
			{
				foreach (XmlAttribute attribute in node.Attributes)
				{
					string callback = GenerateCallback(attribute);
					if (callback is null || callback.Length == 0)
					{
						continue;
					}

					string controlID = GetControlID(node);
					result.Append("<green>'Callback for ");
					if (controlID is null || controlID.Length == 0)
					{
						result.Append(node.Name + "." + attribute.Name);
					}
					else
					{
						result.Append(controlID + " " + attribute.Name);
					}

					result.Append('\n');
					result.Append(callback);
					result.Append("\n\n");
				}
			}

			if (node.HasChildNodes)
			{
				foreach (XmlNode child in node.ChildNodes)
				{
					GenerateCallback(child, result);
				}
			}
			return;
		}

		private static StringBuilder ColorizingCallbacks(StringBuilder callbacks)
		{
			callbacks.Replace("\n", "\\par ");
			callbacks.Replace("<green>", Strings.rtfComment);
			callbacks.Replace("<blue>", Strings.rtfAttributeValue);
			callbacks.Replace("<black>", Strings.rtfAttributeQuote);
			return callbacks;
		}

		private static string GetControlID(XmlNode node)
		{
			if (node.NodeType != XmlNodeType.Element)
			{
				return null;
			}
			try
			{
				foreach (XmlAttribute attribute in node.Attributes)
				{
					if (attribute.Name == "id" || attribute.Name == "idMso" || attribute.Name == "idQ")
					{
						return attribute.Value.Substring(attribute.Value.LastIndexOf(':') + 1);
					}
				}
			}
			catch (XmlException ex)
			{
				System.Diagnostics.Debug.Assert(false, ex.Message);
			}
			return null;
		}

		private static string GenerateCallback(XmlAttribute callback)
		{
			if (callback.Value is null || callback.Value.Length == 0)
			{
				return string.Empty;
			}

			string callbackValue = callback.Value.Substring(callback.Value.LastIndexOf('.') + 1);

			if (attributeList.ContainsKey(callbackValue))
			{
				return string.Empty;
			}

			attributeList.Add(callbackValue, callbackValue);
			BaseCallbackType callbackType = MapToBase(callback);

			switch (callbackType)
			{
				case BaseCallbackType.buttonOnAction: return GenerateVoidCallback(callbackValue);
				case BaseCallbackType.commandOnAction: return GenerateVoidCommand(callbackValue);
				case BaseCallbackType.toggleButtonOnAction: return GenerateToggleButtonOnActionCallback(callbackValue);
				case BaseCallbackType.galleryOnAction: return GenerateItemVoidCallback(callbackValue);
				case BaseCallbackType.comboBoxOnChange:	return GenerateVoidOnChangeCallback(callbackValue);
				case BaseCallbackType.getBoolean:
				case BaseCallbackType.getString:
				case BaseCallbackType.getInt:
				case BaseCallbackType.getImage:
				case BaseCallbackType.getSize: return GenerateReturnCallback(callbackValue);
				case BaseCallbackType.getItemString:
				case BaseCallbackType.getItemImage: return GenerateItemReturnCallback(callbackValue);
				case BaseCallbackType.onLoad: return GenerateOnLoadCallback(callbackValue);
				case BaseCallbackType.onShow: return GenerateOnShowCallback(callbackValue);
				case BaseCallbackType.loadImage: return GenerateLoadImageCallback(callbackValue);
				case BaseCallbackType.getStyle: return GenerateGetSlabStyleCallback(callbackValue);
				case BaseCallbackType.unKnown: return string.Empty;
				default: return string.Empty;
			}
		}

		private const string CONTROL_STRING = "control <blue>As<black> IRibbonControl";
		private const string SUB = "<blue>Sub<black> ";
		private const string ENDSUB = "\n<blue>End Sub<black>";

		private static string GenerateOnLoadCallback(string callback)
		{
			return SUB + callback + "(ribbon <blue>As<black> IRibbonUI)" + ENDSUB;
		}

		private static string GenerateOnShowCallback(string callback)
		{
			return SUB + callback + "(contextObject <blue>As Object<black>)" + ENDSUB;
		}

		private static string GenerateLoadImageCallback(string callback)
		{
			return SUB + callback + "(imageID <blue>As String<black>, <blue>ByRef<black> returnedVal)" + ENDSUB;
		}

		private static string GenerateVoidCommand(string callback)
		{
			return SUB + callback + "(" + CONTROL_STRING + ", <blue>ByRef<black> cancelDefault)" + ENDSUB;
		}

		private static string GenerateVoidCallback(string callback)
		{
			return SUB + callback + "(" + CONTROL_STRING + ")" + ENDSUB;
		}

		private static string GenerateVoidOnChangeCallback(string callback)
		{
			return SUB + callback + "(" + CONTROL_STRING + ", text <blue>As String<black>)" + ENDSUB;
		}

		private static string GenerateToggleButtonOnActionCallback(string callback)
		{
			return SUB + callback + "(" + CONTROL_STRING + ", pressed <blue>As Boolean<black>)" + ENDSUB;
		}

		private static string GenerateReturnCallback(string callback)
		{
			return SUB + callback + "(" + CONTROL_STRING + ", <blue>ByRef<black> returnedVal)" + ENDSUB;
		}

		private static string GenerateItemVoidCallback(string callback)
		{
			return SUB + callback + "(" + CONTROL_STRING + ", id <blue>As String<black>, index <blue>As Integer<black>)" + ENDSUB;
		}

		private static string GenerateItemReturnCallback(string callback)
		{
			return SUB + callback + "(" + CONTROL_STRING + ", index <blue>As Integer<black>, <blue>ByRef<black> returnedVal)" + ENDSUB;
		}

		private static string GenerateGetSlabStyleCallback(string callback)
		{
			return SUB + callback + "(" + CONTROL_STRING + ", <blue>ByRef<black> returnedVal)" + "\n\treturnedVal = BackstageGroupStyle.BackstageGroupStyleWarning" + ENDSUB;
		}

		private static BaseCallbackType MapToBase(XmlAttribute callback)
		{
			switch (callback.Name)
			{
				case "onLoad": return BaseCallbackType.onLoad;
				case "onShow":
				case "onHide": return BaseCallbackType.onShow;
				case "loadImage": return BaseCallbackType.loadImage;
				case "onAction":
					switch (callback.OwnerElement.Name)
					{
						case "dropDown":
						case "gallery": return BaseCallbackType.galleryOnAction;
						case "command": return BaseCallbackType.commandOnAction;
						case "toggleButton":
						case "checkBox": return BaseCallbackType.toggleButtonOnAction;
						default: return BaseCallbackType.buttonOnAction;
					}
				case "onChange": return BaseCallbackType.comboBoxOnChange;
				case "getEnabled":
				case "getVisible":
				case "getPressed":
				case "getShowLabel":
				case "getShowImage": return BaseCallbackType.getBoolean;
				case "getLabel":
				case "getScreentip":
				case "getSupertip":
				case "getDescription":
				case "getKeytip":
				case "getSelectedItemId":
				case "getImageMso":
				case "getContent":
				case "getText":
				case "getTitle":
				case "getTarget":
				case "getHelperText": return BaseCallbackType.getString;
				case "getItemLabel":
				case "getItemTooltip":
				case "getItemId": return BaseCallbackType.getItemString;
				case "getImage": return BaseCallbackType.getImage;
				case "getItemImage": return BaseCallbackType.getItemImage;
				case "getItemCount":
				case "getItemIndex":
				case "getItemHeight":
				case "getItemWidth":
				case "getSelectedItemIndex": return BaseCallbackType.getInt;
				case "getSize":
				case "getItemSize": return BaseCallbackType.getSize;
				case "getStyle": return BaseCallbackType.getStyle;
				default:
					if (callback.Name.StartsWith("on") || callback.Name.StartsWith("get"))
					{
						return BaseCallbackType.unKnown;
					}
					else
					{
						return BaseCallbackType.notCallback;
					}
			}
		}

		private static System.Collections.Hashtable attributeList;

		private enum BaseCallbackType
		{
			buttonOnAction,
			toggleButtonOnAction,
			commandOnAction,
			galleryOnAction,
			comboBoxOnChange,
			getBoolean,
			getString,
			getItemString,
			getImage,
			getItemImage,
			getInt,
			getSize,
			getStyle,
			onShow,
			onLoad,
			loadImage,
			unKnown,
			notCallback, //Always last
		}
	}
}
