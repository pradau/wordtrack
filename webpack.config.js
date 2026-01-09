const path = require('path');
const fs = require('fs');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (env, argv) => {
  const isProduction = argv.mode === 'production';

  return {
    entry: {
      taskpane: './src/taskpane/taskpane.ts',
      commands: './src/commands/commands.ts'
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].js',
      clean: true
    },
    resolve: {
      extensions: ['.ts', '.js']
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          use: 'ts-loader',
          exclude: /node_modules/
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        }
      ]
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'taskpane.html',
        chunks: ['taskpane']
      }),
      new HtmlWebpackPlugin({
        template: './src/commands/commands.html',
        filename: 'commands.html',
        chunks: ['commands']
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: 'manifest.xml', to: 'manifest.xml' }
        ]
      })
    ],
    externals: {
      'office-js': 'Office'
    },
    devServer: {
      static: {
        directory: path.join(__dirname, 'dist')
      },
      port: 3000,
      hot: true,
      client: {
        overlay: false
      },
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      server: (() => {
        // Try to use office-addin-dev-certs for HTTPS
        const certPath = path.join(process.env.HOME || process.env.USERPROFILE || '', '.office-addin-dev-certs', 'localhost.crt');
        const keyPath = path.join(process.env.HOME || process.env.USERPROFILE || '', '.office-addin-dev-certs', 'localhost.key');
        
        if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
          return {
            type: 'https',
            options: {
              key: fs.readFileSync(keyPath),
              cert: fs.readFileSync(certPath)
            }
          };
        } else {
          // Fallback to HTTP if certificates not found
          console.warn('HTTPS certificates not found, using HTTP (may have mixed content issues)');
          console.warn('To fix: Run "npx office-addin-dev-certs install" first');
          return 'http';
        }
      })()
    },
    devtool: isProduction ? false : 'source-map'
  };
};

