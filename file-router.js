const path = require("path");
const fs = require("fs");


const WORD_EXTENSIONS = new Set([
    ".doc",
    ".docx",
    ".odt",
    ".rtf"
]);


const EXCEL_EXTENSIONS = new Set([
    ".xls",
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".csv",
    ".ods"
]);


const POWERPOINT_EXTENSIONS = new Set([
    ".ppt",
    ".pptx",
    ".pps",
    ".ppsx",
    ".odp"
]);


function cleanArgument(argument) {

    if (!argument) {
        return null;
    }


    if (
        argument.startsWith(
            "file://"
        )
    ) {

        try {

            return decodeURIComponent(
                new URL(argument)
                    .pathname
            );

        } catch (_) {

            return argument
                .replace(
                    /^file:\/\//,
                    ""
                );
        }
    }


    return argument;
}


function detectService(filePath) {

    const extension =
        path.extname(
            filePath
        )
        .toLowerCase();


    if (
        WORD_EXTENSIONS.has(
            extension
        )
    ) {

        return "word";
    }


    if (
        EXCEL_EXTENSIONS.has(
            extension
        )
    ) {

        return "excel";
    }


    if (
        POWERPOINT_EXTENSIONS.has(
            extension
        )
    ) {

        return "powerpoint";
    }


    return null;
}


function findOfficeFiles(argumentsList) {

    const files = [];


    for (
        const argument
        of argumentsList
    ) {

        const candidate =
            cleanArgument(
                argument
            );


        if (!candidate) {
            continue;
        }


        if (
            candidate ===
            "--background"
        ) {
            continue;
        }


        try {

            if (
                fs.existsSync(
                    candidate
                ) &&
                fs.statSync(
                    candidate
                )
                .isFile()
            ) {

                const service =
                    detectService(
                        candidate
                    );


                if (service) {

                    files.push({
                        path:
                            path.resolve(
                                candidate
                            ),

                        name:
                            path.basename(
                                candidate
                            ),

                        extension:
                            path.extname(
                                candidate
                            )
                            .toLowerCase(),

                        service
                    });
                }
            }

        } catch (_) {}
    }


    return files;
}


module.exports = {
    detectService,
    findOfficeFiles
};
