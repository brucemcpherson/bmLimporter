
class Cacher {
  constructor({ cacheStore, cacheSeconds, prefix = 'bm-cacher-' }) {
    this.crusher = cacheStore ? new bmCrusher.CrusherPluginCacheService().init({
      store: cacheStore,
      prefix
    }) : null
  }
  get(key) {
    const cached = this.crusher && this.crusher.get(key)
    if (cached) {
      cached.cached = true
    }
    return cached
  }
  put(key, data, cacheSeconds) {
    return this.crusher ? this.crusher.put(key, data, cacheSeconds || this.cacheSeconds) : null
  }
  remove(key) {
    return this.crusher ? this.crusher.remove(key) : null
  }

}