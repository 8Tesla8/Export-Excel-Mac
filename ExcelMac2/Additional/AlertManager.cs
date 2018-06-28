
namespace ExcelMac2 {
	internal interface IAlertWindow {
		bool ShowAlert(string title, string information, string okBtnText);
		bool ShowConfirm(string title, string information, string okBtnText, string cancelBtnText);

		//for spider 3.0
		//ConfirmResult ShowAlert(string message,
								//ConfirmButtonsType buttonsType,
								//ConfirmResult focusButton = ConfirmResult.None,
								//string header = null);
	}


	internal static class AlertManager {
		private readonly static IAlertWindow _alertWindow;

		static AlertManager() {
            _alertWindow = new AlertWindow();
		}


		////for spider 3.0
		//public static ConfirmResult ShowAlertWindow(string message,
		//							   ConfirmButtonsType buttonsType,
		//							   ConfirmResult focusButton = ConfirmResult.None,
		//							   string header = null) {

		//	return _alertWindow.ShowAlert(message, buttonsType, focusButton, header);
		//}


		public static bool ShowAlert(string title, string information, string okBtnText) {
			return _alertWindow.ShowAlert(title, information, okBtnText);
		}


		public static bool ShowConfirm(string title, string information, string okBtnText, string cancelBtnText) {
			return _alertWindow.ShowConfirm(title, information, okBtnText, cancelBtnText);
		}
	}
}
