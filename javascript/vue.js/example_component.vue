<template>
  <div class="authentication">
    <div v-if="loading">
      <BlockUI :message="blockuimsg"></BlockUI>
    </div>

    <div class="container">
      <div class="row">
        <div class="col-lg-4 col-sm-12">
          <form class="card auth_form" @submit="login">
              <div class="header">
                <img class="logo" src="@/assets/images/logo.svg" alt="">
                <h5>Login</h5>
              </div>
              <div class="body">
                <div class="input-group mb-3">
                  <input type="email" class="form-control" placeholder="Email"
                    v-model="email">
                  <div class="input-group-append">
                      <span class="input-group-text"><i class="zmdi zmdi-account-circle"></i></span>
                  </div>
                </div>
                <div class="input-group mb-3">
                  <input type="password" class="form-control" placeholder="Password"
                    v-model="password">
                  <div class="input-group-append">
                    <span class="input-group-text"><i class="zmdi zmdi-lock"></i></span>
                  </div>
                </div>
                <div class="checkbox">
                  <input id="remember_me" type="checkbox">
                  <label for="remember_me">Recuerdame</label>
                </div>
                <button type="submit"
                  class="btn btn-primary btn-block waves-effect waves-light">ENTRAR</button>
              </div>
          </form>
          <div class="copyright text-center">
              &copy;
              {{ year }}
          </div>
        </div>
        <div class="col-lg-8 col-sm-12">
          <div class="card">
            <img src="@/assets/images/signin.svg" alt="Sign In"/>
          </div>
        </div>
      </div>
    </div>
</div>
</template>

<script>
export default {
  name: 'login',

  data() {
    return {
      email: '',
      password: '',
      year: new Date().getFullYear(),
    };
  },
  methods: {
    async login(e) {
      e.preventDefault();

      try {
        this.loading = true;

        const { email, password } = this;
        const response = await this.axios.post('auth/login', { email, password });
        localStorage.setItem('token', response.data.jwt);

        const responseMe = await this.axios.post('auth/me');
        localStorage.setItem('user', JSON.stringify(responseMe.data));
        this.loading = false;

        this.$toasted.show('Login satisfactorio', { type: 'success' });
        this.$router.push('dashboard/index');
      } catch (error) {
        this.loading = false;
        this.$toasted.show('Ocurrió un error. Intente de nuevo por favor', { type: 'warning' });
      }
    },
  },
};
</script>