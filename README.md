# excel2csvjson
Small tool to convert .xlsx to csv steam (or json)

# Usage

```shell
$ excel2csvjson -h
excel2csvjson 0.1.1
bingxiao
Convert excel to csv

USAGE:
    excel2csvjson [FLAGS] [OPTIONS] <path>

FLAGS:
    -f, --first        process the 1st if multiple sheets were found. Or use '--sheet' to clarify
    -h, --help         Prints help information
    -J, --json         output in JSON instead <experimental>
    -P, --pretty       Enable pretty printing
    -V, --version      Prints version information
    -v, --verbosity    Pass many times for more log output

OPTIONS:
        --sheet <sheet>    Which sheet to convert

ARGS:
    <path>    Which excel to convert
```


# Acknowledgement
Many thanks to all these decent rust libraries:

- [quicli](https://github.com/killercup/quicli) - CLI app builder
- [calamine](https://github.com/tafia/calamine) - Excel file reader
- [csv](https://github.com/BurntSushi/rust-csv) - CSV manipulator
- [serde_json](https://github.com/serde-rs/json) - Strongly typed JSON library

and [Rust](https://www.rust-lang.org/) of course.
