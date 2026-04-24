import { CanActivateFn } from '@angular/router';
import axios from 'axios';
import { environment } from 'src/environments/environment';

export const guardGuard: CanActivateFn = async (route, state) => {
  const token = localStorage.getItem('token');
  const url = environment.urlapi

  if (!token) {
    window.alert('Debes iniciar sesión para continuar.')
    window.location.href = '/login';
    return false
  } else {
    try {
      const user = await axios.get(`${url}/users/me?populate=role`, {
        headers: {
          'Authorization': `Bearer ${token}`
        }
      })
      if (user.data.role.type == 'Admin' || user.data.role.type == 'Gerente') {
        localStorage.removeItem('token')
        localStorage.removeItem('user')
        window.alert('Tu sesión caduco, vuelve a iniciar sesion.')
        window.location.href = '/login';
        return false
      }
      return true
    } catch (error) {
      localStorage.removeItem('token')
      localStorage.removeItem('user')
      window.alert('Tu sesión caduco, vuelve a iniciar sesion.')
      window.location.href = '/login';
      return false
    }
  }
  window.location.href = '/login';
  return false;
};
