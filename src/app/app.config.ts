import { ApplicationConfig } from '@angular/core';
import { provideRouter } from '@angular/router';
import { routes } from './app.routes';
import { provideHttpClient, withInterceptorsFromDi, HTTP_INTERCEPTORS } from '@angular/common/http';
import { 
  MsalGuard, 
  MsalInterceptor, 
  MsalService, 
  MsalBroadcastService, 
  MSAL_INSTANCE, 
  MSAL_GUARD_CONFIG, 
  MSAL_INTERCEPTOR_CONFIG, 
  MsalGuardConfiguration, 
  MsalInterceptorConfiguration 
} from '@azure/msal-angular';
import { 
  IPublicClientApplication, 
  PublicClientApplication, 
  InteractionType, 
  BrowserCacheLocation 
} from '@azure/msal-browser';

export function MSALInstanceFactory(): IPublicClientApplication {
  return new PublicClientApplication({
    auth: {
      clientId: 'TU_CLIENT_ID_AQUI', 
      authority: 'https://login.microsoftonline.com/TU_TENANT_ID_AQUI', 
      redirectUri: 'http://localhost:4200'
    },
    cache: {
      cacheLocation: BrowserCacheLocation.LocalStorage
    }
  });
}

export function MSALInterceptorConfigFactory(): MsalInterceptorConfiguration {
  const protectedResourceMap = new Map<string, Array<string>>();
  protectedResourceMap.set('http://localhost:8080/', ['api://TU_CLIENT_ID_AQUI/access_as_user']);

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap
  };
}

export function MSALGuardConfigFactory(): MsalGuardConfiguration {
  return {
    interactionType: InteractionType.Redirect,
    authRequest: {
      scopes: ['user.read']
    }
  };
}

export const appConfig: ApplicationConfig = {
  providers: [
    provideRouter(routes),
    provideHttpClient(withInterceptorsFromDi()),
    {
      provide: HTTP_INTERCEPTORS,
      useClass: MsalInterceptor,
      multi: true
    },
    {
      provide: MSAL_INSTANCE,
      useFactory: MSALInstanceFactory
    },
    {
      provide: MSAL_GUARD_CONFIG,
      useFactory: MSALGuardConfigFactory
    },
    {
      provide: MSAL_INTERCEPTOR_CONFIG,
      useFactory: MSALInterceptorConfigFactory
    },
    MsalService,
    MsalGuard,
    MsalBroadcastService
  ]
};