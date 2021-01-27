package sample.actionhandler;
import javax.security.auth.Subject;

import com.filenet.api.core.Connection;
import com.filenet.api.core.Factory;
import com.filenet.api.util.UserContext;

import com.ibm.casemgmt.api.context.CaseMgmtContext;

public class ConnectionClass {
	static String uri = "http://ibmbaw:9080/wsi/FNCEWS40MTOM";
	static String username = "dadmin";
	static String password = "dadmin";
	static UserContext old = null;
	static CaseMgmtContext oldCmc = null;

	public Connection getConnection() {
		Connection conn = Factory.Connection.getConnection(uri);
		Subject subject = UserContext.createSubject(conn, username, password, "FileNetP8WSI");
		UserContext.get().pushSubject(subject);
		return conn;
	}

}
