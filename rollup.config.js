// rollup.config.js
import resolve from 'rollup-plugin-node-resolve';

export default {
  input   : 'lib/index.js',
  output: {
  	format: 'umd',
  	name  : 'renderExcelJS',
    file  : 'index.js'
  },
  plugins : [resolve()]
};