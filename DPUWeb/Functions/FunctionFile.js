'use strict';

(function () {
	Office.initialize = function (reason) {

	};
})();


function toggleProtection(args) {
	Excel.run(function(context) {

		const sheet = context.workbook.worksheets.getActiveWorksheet();
		sheet.load('protection/protected');

		return context.sync()
			.then(
				function() {
					if (sheet.protection.protected) {
						sheet.protection.unprotect('taxi');
					} else {
						sheet.protection.protect(null, 'taxi');
					}
				}				
			)
			.then(context.sync);
	})
	.catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
	});
	args.completed();
}