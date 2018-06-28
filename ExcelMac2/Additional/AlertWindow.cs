using System;
using AppKit;

namespace ExcelMac2 {

	internal class AlertWindow : IAlertWindow {

		//for spider 3.0
		//public ConfirmResult ShowAlert(string message,
		//							   ConfirmButtonsType buttonsType,
		//							   ConfirmResult focusButton = ConfirmResult.None,
		//							   string header = null
		//							  ) {

		//	if (buttonsType == ConfirmButtonsType.OK) {
		//		if (header == null)
		//			header = Resources.Confirm_Common_Header_Alert;

		//		ShowAlert(header, message, Resources.All_Button_OK);
		//		return ConfirmResult.None;
		//	}
		//	else if (buttonsType == ConfirmButtonsType.OKCancel) {
		//		if (header == null)
		//			header = Resources.Confirm_Common_Header_Confirm;

		//		var result = ShowConfirm(header, message, Resources.All_Button_OK, Resources.All_Button_Cancel);

		//		if (result)
		//			return ConfirmResult.OK;

		//		return ConfirmResult.Cancel;
		//	}
		//	else if (buttonsType == ConfirmButtonsType.YesNo) {
		//		if (header == null)
		//			header = Resources.Confirm_Common_Header_Confirm;
				
		//		var result = ShowConfirm(header, message, Resources.All_Button_YES, Resources.All_Button_NO);

		//		if (result)
		//			return ConfirmResult.Yes;

		//		return ConfirmResult.No;
		//	}

		//	throw new System.NotImplementedException("do not have actin for " +
		//												 buttonsType.ToString());
		//}
		//

		public bool ShowAlert(string title, string information, string okBtnText) {

			NSApplication.SharedApplication.InvokeOnMainThread(() => {
				var alert = CreateAlert(title, information, okBtnText, null);
				alert.RunModal();
			});

			return false;
		}


		public bool ShowConfirm(string title, string information, string okBtnText, string canselBtnText) {
			nint result = 0;

			NSApplication.SharedApplication.InvokeOnMainThread(() => {
				var alert = CreateAlert(title, information, okBtnText, canselBtnText);
				result = alert.RunModal();
			});

			if (result == 1000)
				return true;

			return false;
		}


		private NSAlert CreateAlert(string title, string information, string okBtnText, string canselBtnText) {
			var alertStyle = NSAlertStyle.Critical;

			if (!string.IsNullOrEmpty(canselBtnText))
				alertStyle = NSAlertStyle.Informational;

			var alert = new NSAlert() {
				MessageText = title,
				InformativeText = information,
				AlertStyle = alertStyle,
			};

			switch (alertStyle) {
				case NSAlertStyle.Critical:
					alert.AddButton(okBtnText);     //1000
					break;

				case NSAlertStyle.Informational:
					alert.AddButton(okBtnText);     //1000
					alert.AddButton(canselBtnText); //1001

					var alert1 = new NSAlert();
					alert1.AddButton("OK");
					alert1.AddButton("Cancel");

					alert.Buttons[1].KeyEquivalent = alert1.Buttons[1].KeyEquivalent;
					break;
			}

			return alert;
		}
	}
}
