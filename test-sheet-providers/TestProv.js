export default {
  name: "Test Provider",
  type: "testing",

  async create(context) {
    return {
      async load(sheet) {
        return 1
      }
    }
  }
}