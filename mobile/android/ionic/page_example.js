import { IonContent, IonToast, IonItem, IonItemDivider, IonList, IonLabel } from '@ionic/react';
import React from 'react';
import Header from '../components/Header';
import { getListRegister } from '../services/ApiService';

const ListRegister = () => {
  const [showToast, setShowToast] = React.useState(false);
  const [registers, setRegisters] = React.useState([]);
  const [dateFrom, setDateFrom] = React.useState(null);

  /**
   * @method fetchData
   * @description Obtiene los movimientos
   */
  async function fetchData(dateTo) {
    if (dateFrom !== null && dateTo !== null) {
      const personalId = window.localStorage.getItem('personal');
      const response = await getListRegister(personalId, dateFrom, dateTo);

      if (response.status === 200) {
        const data = response.data.data;
        setRegisters(data);
      } else {
        setShowToast(true);
      }
    }
  }

  return (
    <>
      <Header />
      <IonContent className="ion-padding">
        <IonToast
          isOpen={showToast}
          onDidDismiss={() => setShowToast(false)}
          message={'Ocurrió un error al leer los datos'}
          duration={400}
        />

        <IonList>
          Desde
          <input type="date" from="datefrom" onChange={(event) => setDateFrom(event.target.value)} />
          Hasta
          <input type="date" name="dateto" onChange={(event) => fetchData(event.target.value)} />
          <br /><br />

          <IonItemDivider color="primary">
            <IonLabel>
              Resultados
            </IonLabel>
          </IonItemDivider>

          {registers && registers.map(row => (
            <IonItem key={row.id}>
              <IonLabel>{new Date(row.datetime).toLocaleDateString()}</IonLabel>
              <IonLabel>{row.movementType === 'E' ? 'Entrada' : 'Salida'}</IonLabel>
              <IonLabel>{row.place} </IonLabel>
            </IonItem>
          ))}
        </IonList>
      </IonContent>
    </>
  );
};

export default ListRegister;
