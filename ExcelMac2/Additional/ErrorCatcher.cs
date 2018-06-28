using System;
using Foundation;

namespace ExcelMac2 {
	internal static class ErrorCatcher {

		public static void CurrentDomainUnhandledException(object sender, UnhandledExceptionEventArgs e) {
			Error(e.ExceptionObject as Exception);
		}


		public static void Error(NSException exception) {
			if (exception != null) {
				AlertManager.ShowAlert("Native Exception", exception.ToString(), "OK");
			}
		}


		public static void Error(Exception exception) {
			if (exception != null) {
				AlertManager.ShowAlert("C# Exception", exception.ToString(), "OK");
			}
		}
	}
}
