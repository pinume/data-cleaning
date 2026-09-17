fn main() {
    if let Err(error) = data_cleaning::app::cli::run() {
        eprintln!("{error}");
        std::process::exit(1);
    }
}
