const path = require('path');
const webpack = require('webpack');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
    entry: {
        main: './js/domino/init.js',
        second: './js/solitaire/init.js',
        third: './js/me.js',
        jenga: './js/jenga/init.js',
    },
    output: {
        filename: '[name].[contenthash].bundle.js',
        path: path.resolve(__dirname, 'dist'),
        clean: true,
    },
    plugins: [
        new HtmlWebpackPlugin({
            template: './src/domino.html',
            filename: 'domino.html',
            inject: 'body',
            scriptLoading: 'blocking',
            chunks: ['main'],
            minify: {
                collapseWhitespace: true,
                keepClosingSlash: true,
                removeComments: true,
                removeRedundantAttributes: false, // do not remove type="text"
                removeScriptTypeAttributes: true,
                removeStyleLinkTypeAttributes: true,
                useShortDoctype: true
            }
        }),
        new HtmlWebpackPlugin({
            template: './src/solitaire.html',
            filename: 'solitaire.html',
            inject: 'body',
            scriptLoading: 'blocking',
            chunks: ['second'],
            minify: {
                collapseWhitespace: true,
                keepClosingSlash: true,
                removeComments: true,
                removeRedundantAttributes: false, // do not remove type="text"
                removeScriptTypeAttributes: true,
                removeStyleLinkTypeAttributes: true,
                useShortDoctype: true
            }
        }),
        new HtmlWebpackPlugin({
            template: './src/index.html',
            filename: 'index.html',
            inject: 'body',
            scriptLoading: 'blocking',
            chunks: ['third'],
            minify: {
                collapseWhitespace: true,
                keepClosingSlash: true,
                removeComments: true,
                removeRedundantAttributes: false, // do not remove type="text"
                removeScriptTypeAttributes: true,
                removeStyleLinkTypeAttributes: true,
                useShortDoctype: true
            }
        }),
        new HtmlWebpackPlugin({
            template: './src/jenga.html',
            filename: 'jenga.html',
            inject: 'body',
            scriptLoading: 'blocking',
            chunks: ['jenga'],
            minify: {
                collapseWhitespace: true,
                keepClosingSlash: true,
                removeComments: true,
                removeRedundantAttributes: false, // do not remove type="text"
                removeScriptTypeAttributes: true,
                removeStyleLinkTypeAttributes: true,
                useShortDoctype: true
            }
        }),
        new CopyWebpackPlugin({
            patterns: [
                { from: 'brand-specific/brand.css', to: 'brand-specific/brand.css' },
                { from: 'assets', to: 'assets' },
                {from: 'css', to: 'css'},
                {from: 'src/service-catalog.csv', to: 'service-catalog.csv'},
                {from: 'src/people-database.csv', to: 'people-database.csv'},
                {from: 'src/jira-cards.csv', to: 'jira-cards.csv'},
                {from: 'src/jenga-events.csv', to: 'jenga-events.csv'},
                {from: 'src/custom-filters.csv', to: 'custom-filters.csv'},
                {from: 'src/robots.txt', to: 'robots.txt'},
                {from: 'src/sitemap.xml', to: 'sitemap.xml'},
                { from: 'assets', to: 'assets' },
                // Ship the games to Pages, but not the Electron desktop project
                // (its source/build artifacts aren't part of the website).
                { from: 'game', to: 'game', globOptions: { ignore: ['**/desktop/**'] } },
            ],
        }),
        new webpack.DefinePlugin({
            __APP_BUILD__: JSON.stringify(process.env.APP_BUILD || 'dev'),
            __BUILD_DATE__: JSON.stringify(process.env.BUILD_DATE || (() => {
                const d = new Date();
                return d.toLocaleString('sv', { timeZone: 'Europe/Rome' }).slice(0, 16).replace('T', ' ') + ' CEST';
            })()),
            __FEATURE_LOD__: JSON.stringify(process.env.FEATURE_LOD !== 'false'),
            __ITSM_RETENTION_DAYS__: JSON.stringify(Number(process.env.ITSM_RETENTION_DAYS || 100)),
        }),
    ],
    module: {
        rules: [
            {
                test: /\.js$/,
                exclude: /node_modules/,
                use: {
                    loader: 'babel-loader',
                },
            },
            {
                test: /\.css$/,
                use: ['style-loader', 'css-loader'],
            },
        ],
    },
};