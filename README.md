# Referencehelper

Helps you to comfortably insert successive references.

## Getting Started

For example: A reference like `25100-300`. The right end of the reference, e.g. `300`, will be continued in steps of 1. The new reference would be `25100-301` and so forth. The same is done for endings with letters, e.g. `25100a`.

## Hotkeys

<kbd>Win</kbd> + <kbd>1</kbd>: Copy highlighted reference.

<kbd>Win</kbd> + <kbd>2</kbd>: Replaces the highlighted text or, if no text is selected, the first character before the mouse pointer with the next reference.

<kbd>Win</kbd> + <kbd>3</kbd>: Resets the next reference by one step.

<kbd>Alt</kbd> + <kbd>Esc</kbd>: Exit script.

## Installing 

* Download the latest release and execute the script.
  * If you don't want to install [Autohotkey](https://www.autohotkey.com/) the ZIP-file from https://autohotkey.com/download/ includes the compiler program to create an executable `.EXE`from the release file `.AHK`.
  * For additional information please refer to https://ahkde.github.io/docs/Tutorial.htm. 

## Using

1. Highlight your reference (with your mouse), e.g. `25100-300`.
1. Press <kbd>Win</kbd> + <kbd>1</kbd> to copy.
1. Select the position where you want the new reference to be inserted. _Please note: if no text is selected, the first character before the mouse pointer will be replaced._
1. Press <kbd>Win</kbd> + <kbd>2</kbd> to insert the new reference, e.g. `25100-301`.
1. Select the position where you want the next reference to be inserted. _Please note: if no text is selected, the first character before the mouse pointer will be replaced._
1. Press <kbd>Win</kbd> + <kbd>2</kbd> to insert the next reference, e.g. `25100-302`.
1. ...

## Built With

* [Autohotkey](https://www.autohotkey.com/) - The keyboard macro program used

## Contributing

Please contact the author.

## Authors

* **Felix L. Wenger** - [felixwenger](https://github.com/felixwenger)

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.
