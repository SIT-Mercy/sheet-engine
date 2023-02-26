export default {
  name: "Test Provider",
  type: "testing",

  async create(context) {
    return {
      async load() {
        return 1
      }
    }
  }
}