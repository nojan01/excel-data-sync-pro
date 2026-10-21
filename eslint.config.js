const js = require('@eslint/js');
const globals = require('globals');

module.exports = [
    {
        ignores: ['node_modules/**', 'dist/**', 'build/**', 'python/**', 'python-embed/**', 'docs/**']
    },
    js.configs.recommended,
    {
        rules: {
            'no-unused-vars': ['warn', { caughtErrors: 'none' }],
            'no-useless-escape': 'warn',
            'no-useless-assignment': 'warn',
            'no-case-declarations': 'warn',
            'no-control-regex': 'warn',
            'no-prototype-builtins': 'warn'
        }
    },
    {
        files: ['**/*.js'],
        languageOptions: {
            ecmaVersion: 'latest',
            sourceType: 'commonjs',
            globals: {
                ...globals.node
            }
        }
    },
    {
        files: ['src/**/*.js'],
        languageOptions: {
            sourceType: 'script',
            globals: {
                ...globals.browser
            }
        }
    },
    {
        files: ['preload.js'],
        languageOptions: {
            globals: {
                ...globals.browser
            }
        }
    }
];
