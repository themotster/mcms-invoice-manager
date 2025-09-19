const path = require('path');

module.exports = {
  target: 'electron-renderer',
  entry: path.resolve(__dirname, 'renderer', 'App.jsx'),
  output: {
    filename: 'bundle.js',
    path: path.resolve(__dirname, 'renderer')
  },
  resolve: {
    extensions: ['.js', '.jsx']
  },
  module: {
    rules: [
      {
        test: /\.jsx?$/,
        exclude: /node_modules/,
        use: {
          loader: 'babel-loader',
          options: {
            presets: ['@babel/preset-env', '@babel/preset-react']
          }
        }
      }
    ]
  },
  devtool: 'source-map'
};
