import jsbeautifier, sys

activate_file = "../../myapp/Scripts/activate_this.py"
execfile(activate_file, dict(__file__=activate_file))

code = sys.argv[1]
opts = jsbeautifier.default_options()
opts.indent_char = "&nbsp"
print jsbeautifier.beautify(code, opts)
