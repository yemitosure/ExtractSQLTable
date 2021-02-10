
    module.exports = {
        DatabaseConfig: {
            server: "localhost\\sqlexpress",
            port: "1433",
            database: "master",
            driver: "msnodesqlv8",
            options: {
                trustedConnection: true
            }
        },
        SQLStatement: "select * from sysobjects",
        FilePath:"Test.xlsx"
    };