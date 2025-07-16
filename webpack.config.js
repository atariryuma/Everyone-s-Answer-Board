const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const GasPlugin = require('gas-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
  mode: 'development',
  entry: {
    main: './src/main.gs', // Your main GAS file
    'dev-styles': './src/dev-entry.js', // For local development styles
    // Add other entry points if you have multiple independent GAS files
  },
  output: {
    filename: '[name].js',
    path: path.resolve(__dirname, 'dist'),
    clean: true,
  },
  resolve: {
    extensions: ['.js', '.gs'],
  },
  module: {
    rules: [
      {
        test: /\.html$/,
        use: {
          loader: 'html-loader',
          options: {
            sources: false,
          },
        },
      },
      {
        test: /\.gs$/,
        exclude: /node_modules/,
        loader: 'babel-loader', // You might need babel-loader for modern JS features
        options: {
          presets: ['@babel/preset-env'],
        },
      },
      {
        test: /\.css$/,
        use: [
          'style-loader',
          'css-loader',
          {
            loader: 'postcss-loader',
            options: {
              postcssOptions: {
                plugins: [
                  require('tailwindcss'),
                  require('autoprefixer'),
                ],
              },
            },
          },
        ],
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: './src/LoginPage.html', // Your main HTML file for the web app
      filename: 'index.html',
      inject: 'body',
      chunks: ['main', 'dev-styles'], // Ensure this matches your entry point
    }),
    new GasPlugin(),
    new CopyWebpackPlugin({
      patterns: [
        { from: 'src/**/*.html', to: '[name][ext]', context: 'src/', noErrorOnMissing: true },
        { from: 'src/**/*.gs', to: '[name][ext]', context: 'src/', noErrorOnMissing: true },
        { from: 'src/**/*.js', to: '[name][ext]', context: 'src/', noErrorOnMissing: true },
        { from: 'src/**/*.json', to: '[name][ext]', context: 'src/', noErrorOnMissing: true },
      ],
    }),
  ],
  devServer: {
    static: {
      directory: path.join(__dirname, 'dist'),
    },
    compress: true,
    port: 9000,
    open: true,
  },
};
