library(Rs3tools)

# download the workbook from AWS S3
s3_bucket    <- "alpha-forward-look"
load_path <- Rs3tools::list_files_in_buckets(s3_bucket, path_only = TRUE)
download_file_from_s3(load_path, "Forward Look/Forward Look.xlsx", overwrite=TRUE)