module.exports = {
  testEnvironment: 'node',
  roots: ['<rootDir>'],
  testPathIgnorePatterns: ['/node_modules/', '/release/'],
  moduleFileExtensions: ['js', 'jsx'],
  transform: {
    '^.+\\.[tj]sx?$': 'babel-jest'
  }
};
