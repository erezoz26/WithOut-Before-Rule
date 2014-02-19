package mavenTest;

import org.junit.rules.TestRule;
import org.junit.runner.Description;
import org.junit.runners.model.Statement;

public class WithoutBeforeRule implements TestRule{				

	public Statement apply(final Statement base, final Description description) {
		  return new Statement() {
		    @Override
		    public void evaluate() throws Throwable {
		     WithoutBefore annotation = description.getAnnotation(WithoutBefore.class);		     
		     if ( annotation == null){
		        // put the before method here
		    	 System.out.println("i'm before test");
		      }

		     base.evaluate();
		      
		    }
		  };
		}

	
}
