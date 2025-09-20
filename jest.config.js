module.exports = {
  testEnvironment: 'node',
  roots: ['<rootDir>'],
  moduleFileExtensions: ['js', 'jsx'],
  transform: {
    '^.+\\.[tj]sx?$': 'babel-jest'
  }
};
